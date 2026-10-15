from django import forms
from django import forms
from .models import Stock, WeldingNeed, DocumentationRecord


class ExcelUploadForm(forms.Form):
    FILE_CHOICES = [
        ("stock", "Заготовка"),
        ("welding", "Потребность сварки"),
        ("documentation", "Документация"),
    ]

    data_type = forms.ChoiceField(
        label="Тип данных",
        choices=FILE_CHOICES,
    )

    file = forms.FileField(
        label="Excel-файл",
        widget=forms.ClearableFileInput(
            attrs={"accept": ".xlsx,.xls"}
        ),
    )


class StockEditForm(forms.ModelForm):
    class Meta:
        model = Stock
        fields = [
            "date", "batch_id", "employee", "plan_qty", "fact_qty"
        ]
        widgets = {
            "date": forms.DateInput(attrs={"type": "date"}),
        }


class WeldingNeedEditForm(forms.ModelForm):
    class Meta:
        model = WeldingNeed
        fields = ["date", "batch_id", "need_qty"]
        widgets = {
            "date": forms.DateInput(attrs={"type": "date"}),
        }


class DocumentationEditForm(forms.ModelForm):
    class Meta:
        model = DocumentationRecord
        fields = ["date", "employee", "plan_docs", "fact_docs"]
        widgets = {
            "date": forms.DateInput(attrs={"type": "date"}),
        }
