"""riverdale_admin URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/3.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path
from studyingstatus import views
from django.conf import settings
from django.conf.urls.static import static


urlpatterns = [

    path('', views.home, name='home'),
    path('upload/', views.upload_file, name='upload_file'),
    # Adjusted to remove the filename parameter from the path, as it's now stored in session
    path('download/', views.download_file, name='download_file'),
]