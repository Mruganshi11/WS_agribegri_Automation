import subprocess
import os
import psutil
import json
from django.shortcuts import render, redirect
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.models import User
from django.contrib.auth.decorators import login_required
from django.contrib import messages

# Absolute Paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_FILE = os.path.join(BASE_DIR, 'automation.log')
PID_FILE = os.path.join(BASE_DIR, 'script.pid')
BALANCE_FILE = os.path.join(BASE_DIR, 'balance.txt')
ORDERS_COUNT_FILE = os.path.join(BASE_DIR, 'orders_count.txt')
HISTORY_FILE = os.path.join(BASE_DIR, 'processed_orders.json')
SCRIPT_PATH = os.path.join(BASE_DIR, 'm_happy_flow_final.py')
PYTHON_EXE = os.path.join(BASE_DIR, 'venv', 'Scripts', 'python.exe')

@login_required(login_url='login')
def index(request):
    return render(request, 'dashboard/index.html')

def get_balance_value():
    if not os.path.exists(BALANCE_FILE):
        return 0.0
    try:
        with open(BALANCE_FILE, 'r') as f:
            return float(f.read().strip())
    except:
        return 0.0

def get_orders_count_value():
    if not os.path.exists(ORDERS_COUNT_FILE):
        return 0
    try:
        with open(ORDERS_COUNT_FILE, 'r') as f:
            return int(f.read().strip())
    except:
        return 0

@login_required(login_url='login')
@csrf_exempt
def set_balance(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            amount = float(data.get('amount', 0))
            with open(BALANCE_FILE, 'w') as f:
                f.write(str(round(amount, 2)))
            return JsonResponse({'status': 'success', 'balance': amount})
        except Exception as e:
            return JsonResponse({'status': 'error', 'message': str(e)})
    return JsonResponse({'status': 'error', 'message': 'POST required'})

@login_required(login_url='login')
@csrf_exempt
def run_script(request):
    if request.method == 'POST':
        if not os.path.exists(SCRIPT_PATH):
            return JsonResponse({'status': 'error', 'message': 'Script not found.'}, status=404)
        
        try:
            data = json.loads(request.body)
            username = data.get('username')
            password = data.get('password')
            otp = data.get('otp')
            target_order_id = data.get('targetOrderId', '').strip()
            
            if not all([username, password, otp]):
                return JsonResponse({'status': 'error', 'message': 'Username, Password, and OTP are required.'})
                
            # Check initial balance
            if get_balance_value() < 0.75:
                return JsonResponse({'status': 'error', 'message': 'Insufficient balance. Please recharge.'})

            # Stop existing if any
            stop_script(request)

            with open(LOG_FILE, 'w', encoding='utf-8') as f:
                f.write("--- Automation Started via Dashboard ---\n")
            
            log_f = open(LOG_FILE, 'a', encoding='utf-8')
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'

            cmd = [PYTHON_EXE, '-u', SCRIPT_PATH, username, password, otp]
            if target_order_id:
                cmd.append(target_order_id)

            process = subprocess.Popen(
                cmd,
                cwd=BASE_DIR,
                stdout=log_f,
                stderr=subprocess.STDOUT,
                env=env,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            
            with open(PID_FILE, 'w') as f:
                f.write(str(process.pid))
            
            return JsonResponse({'status': 'success', 'message': f'Started (PID: {process.pid})'})
        except Exception as e:
            return JsonResponse({'status': 'error', 'message': str(e)}, status=500)
            
    return JsonResponse({'status': 'error', 'message': 'POST required'}, status=405)

@login_required(login_url='login')
@csrf_exempt
def stop_script(request):
    pid = None
    if os.path.exists(PID_FILE):
        with open(PID_FILE, 'r') as f:
            try:
                pid = int(f.read().strip())
            except:
                pass
    
    if pid:
        try:
            parent = psutil.Process(pid)
            for child in parent.children(recursive=True):
                child.kill()
            parent.kill()
            if os.path.exists(PID_FILE):
                os.remove(PID_FILE)
            return JsonResponse({'status': 'success', 'message': 'Process stopped.'})
        except:
            if os.path.exists(PID_FILE):
                os.remove(PID_FILE)
    
    return JsonResponse({'status': 'error', 'message': 'No active process to stop.'})

@login_required(login_url='login')
def check_status(request):
    pid = None
    balance = get_balance_value()
    orders_count = get_orders_count_value()
    if os.path.exists(PID_FILE):
        with open(PID_FILE, 'r') as f:
            try:
                pid = int(f.read().strip())
            except:
                pass
    
    is_running = False
    msg = "Offline"
    if pid:
        try:
            process = psutil.Process(pid)
            if process.is_running() and process.status() != psutil.STATUS_ZOMBIE:
                is_running = True
                msg = f'Running (PID: {pid})'
        except:
            if os.path.exists(PID_FILE):
                os.remove(PID_FILE)
    
    return JsonResponse({'running': is_running, 'message': msg, 'balance': balance, 'orders_count': orders_count})

@login_required(login_url='login')
def get_logs(request):
    if not os.path.exists(LOG_FILE):
        return JsonResponse({'logs': 'Log file not found.', 'balance': get_balance_value(), 'orders_count': get_orders_count_value()})
    
    try:
        with open(LOG_FILE, 'r', encoding='utf-8', errors='replace') as f:
            lines = f.readlines()
            return JsonResponse({'logs': "".join(lines[-200:]), 'balance': get_balance_value(), 'orders_count': get_orders_count_value()})
    except Exception as e:
        return JsonResponse({'logs': f'Error: {str(e)}', 'balance': get_balance_value(), 'orders_count': get_orders_count_value()})

@login_required(login_url='login')
def order_history(request):
    history = []
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r') as f:
                history = json.load(f)
                # Show latest first
                history.reverse()
        except:
            history = []
    return render(request, 'dashboard/history.html', {'history': history})

@login_required(login_url='login')
@csrf_exempt
def clear_history_view(request):
    try:
        # Clear JSON history
        if os.path.exists(HISTORY_FILE):
            os.remove(HISTORY_FILE)
        
        # Reset orders count file
        if os.path.exists(ORDERS_COUNT_FILE):
            with open(ORDERS_COUNT_FILE, 'w') as f:
                f.write("0")
                
        return JsonResponse({'status': 'success', 'message': 'History and order count cleared.'})
    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

def signup_view(request):
    if request.method == 'POST':
        email = request.POST.get('email')
        username = request.POST.get('username')
        password = request.POST.get('password')
        confirm_password = request.POST.get('confirm_password')

        if not email or not username or not password or not confirm_password:
            messages.error(request, 'All fields are required.')
            return redirect('signup')

        if password != confirm_password:
            messages.error(request, 'Passwords do not match.')
            return redirect('signup')

        if User.objects.filter(username=username).exists():
            messages.error(request, 'Username already exists.')
            return redirect('signup')
            
        if User.objects.filter(email=email).exists():
            messages.error(request, 'Email already registered.')
            return redirect('signup')

        user = User.objects.create_user(username=username, email=email, password=password)
        user.save()
        messages.success(request, 'Account created successfully. Please login.')
        return redirect('login')

    return render(request, 'dashboard/signup.html')

def login_view(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')

        user = authenticate(request, username=username, password=password)
        if user is not None:
            login(request, user)
            return redirect('index')
        else:
            messages.error(request, 'Invalid username or password.')
            return redirect('login')

    return render(request, 'dashboard/login.html')

def logout_view(request):
    logout(request)
    return redirect('login')
