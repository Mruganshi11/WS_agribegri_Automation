from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('login/', views.login_view, name='login'),
    path('signup/', views.signup_view, name='signup'),
    path('logout/', views.logout_view, name='logout'),
    path('run-script/', views.run_script, name='run_script'),
    path('check-status/', views.check_status, name='check_status'),
    path('stop-script/', views.stop_script, name='stop_script'),
    path('get-logs/', views.get_logs, name='get_logs'),
    path('set-balance/', views.set_balance, name='set_balance'),
    path('history/', views.order_history, name='order_history'),
    path('clear-history/', views.clear_history_view, name='clear_history'),
    # Clubbing Flow
    path('clubbing/', views.clubbing_index, name='clubbing_index'),
    path('clubbing/run/', views.run_clubbing_script, name='run_clubbing_script'),
    path('clubbing/stop/', views.stop_clubbing_script, name='stop_clubbing_script'),
    path('clubbing/status/', views.check_clubbing_status, name='check_clubbing_status'),
    path('clubbing/logs/', views.get_clubbing_logs, name='get_clubbing_logs'),
    path('clubbing/history/', views.clubbing_history, name='clubbing_history'),
    path('clubbing/clear-history/', views.clear_clubbing_history, name='clear_clubbing_history'),
]
