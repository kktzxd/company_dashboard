from django.urls import path
from . import views

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("upload/", views.upload_excel, name="upload_excel"),
    path("export/excel/", views.export_excel, name="export_excel"),
    path("export/csv/", views.export_csv, name="export_csv"),
    path(
        "edit/<str:data_type>/<int:pk>/",
        views.edit_record,
        name="edit_record",
    ),
]
