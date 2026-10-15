from django.db import models


#Заготовка
class Stock(models.Model):
    record_id = models.CharField(max_length=20)
    date = models.DateField()
    batch_id = models.CharField(max_length=20)
    employee = models.CharField(max_length=100)
    plan_qty = models.IntegerField()
    fact_qty = models.IntegerField(null=True, blank=True)

    class Meta:
        ordering = ["date", "record_id"]

    def __str__(self):
        return f"{self.record_id} — {self.batch_id}"


#Потребность сварки
class WeldingNeed(models.Model):
    record_id = models.CharField(max_length=20)
    date = models.DateField()
    batch_id = models.CharField(max_length=20)
    need_qty = models.IntegerField()

    class Meta:
        ordering = ["date", "record_id"]

    def __str__(self):
        return f"{self.record_id} — {self.batch_id}"


#Документация
class DocumentationRecord(models.Model):
    record_id = models.CharField(max_length=20)
    date = models.DateField()
    employee = models.CharField(max_length=100)
    plan_docs = models.IntegerField()
    fact_docs = models.IntegerField()

    class Meta:
        ordering = ["date", "record_id"]

    def __str__(self):
        return self.record_id
