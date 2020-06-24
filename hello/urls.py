from django.urls import path
from hello import views

app_name = "hello"

urlpatterns = [
    path("", views.home, name="home"),
]