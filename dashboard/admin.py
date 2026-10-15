from django.contrib import admin
from .models import Stock, WeldingNeed, DocumentationRecord


@admin.register(Stock)
class BlankRecordAdmin(admin.ModelAdmin):
    list_display = (
        "record_id",
        "date",
        "batch_id",
        "employee",
        "plan_qty",
        "fact_qty",
    )
    list_filter = ("date", "employee")
    search_fields = ("record_id", "batch_id", "employee")


@admin.register(WeldingNeed)
class WeldingNeedAdmin(admin.ModelAdmin):
    list_display = ("record_id", "date", "batch_id", "need_qty")
    list_filter = ("date",)
    search_fields = ("record_id", "batch_id")


@admin.register(DocumentationRecord)
class DocumentationRecordAdmin(admin.ModelAdmin):
    list_display = (
        "record_id",
        "date",
        "employee",
        "plan_docs",
        "fact_docs",
    )
    list_filter = ("date", "employee")
    search_fields = ("record_id", "employee")
