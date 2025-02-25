from django.urls import path
from . import views

urlpatterns = [
    path('api/register/', views.register_user, name='register_user'),
    path('api/save_test_result/', views.save_test_result, name='save_test_result'),
    path('api/login/', views.login_user, name='login_user'),
]