from django.contrib import admin
from django.http import HttpResponse
import csv
import openpyxl
from .models import User, TestNSI, TestResult


@admin.register(User)
class UserAdmin(admin.ModelAdmin):
    list_display = ('username', 'age', 'education', 'speciality')
    actions = ['export_as_csv', 'export_as_xlsx']

    def export_as_csv(self, request, queryset):
        meta = self.model._meta
        field_names = [field.name for field in meta.fields]
        response = HttpResponse(content_type='text/csv')
        response['Content-Disposition'] = f'attachment; filename={meta}.csv'
        writer = csv.writer(response)
        writer.writerow(field_names)
        for obj in queryset:
            writer.writerow([getattr(obj, field) for field in field_names])
        return response

    def export_as_xlsx(self, request, queryset):
        meta = self.model._meta
        field_names = [field.name for field in meta.fields]
        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename={meta}.xlsx'
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(field_names)
        for obj in queryset:
            ws.append([getattr(obj, field) for field in field_names])
        wb.save(response)
        return response

    export_as_csv.short_description = "Export Selected as CSV"
    export_as_xlsx.short_description = "Export Selected as XLSX"


@admin.register(TestNSI)
class TestNSIAdmin(admin.ModelAdmin):
    list_display = ('test_name', 'title_all', 'title_correct')
    actions = ['export_as_csv', 'export_as_xlsx']

    def export_as_csv(self, request, queryset):
        meta = self.model._meta
        field_names = [field.name for field in meta.fields]
        response = HttpResponse(content_type='text/csv')
        response['Content-Disposition'] = f'attachment; filename={meta}.csv'
        writer = csv.writer(response)
        writer.writerow(field_names)
        for obj in queryset:
            writer.writerow([getattr(obj, field) for field in field_names])
        return response

    def export_as_xlsx(self, request, queryset):
        meta = self.model._meta
        field_names = [field.name for field in meta.fields]
        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename={meta}.xlsx'
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(field_names)
        for obj in queryset:
            ws.append([getattr(obj, field) for field in field_names])
        wb.save(response)
        return response

    export_as_csv.short_description = "Export Selected as CSV"
    export_as_xlsx.short_description = "Export Selected as XLSX"


@admin.register(TestResult)
class TestResultAdmin(admin.ModelAdmin):
    list_display = ('user', 'test', 'try_number', 'accuracy')
    actions = ['export_as_csv', 'export_as_xlsx']

    def export_as_csv(self, request, queryset):
        meta = self.model._meta
        field_names = [field.name for field in meta.fields]
        response = HttpResponse(content_type='text/csv')
        response['Content-Disposition'] = f'attachment; filename={meta}.csv'
        writer = csv.writer(response)
        writer.writerow(field_names)
        for obj in queryset:
            writer.writerow([getattr(obj, field) for field in field_names])
        return response

    def export_as_xlsx(self, request, queryset):
        meta = self.model._meta
        field_names = [field.name for field in meta.fields]
        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename={meta}.xlsx'
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(field_names)
        for obj in queryset:
            ws.append([getattr(obj, field) for field in field_names])
        wb.save(response)
        return response

    export_as_csv.short_description = "Export Selected as CSV"
    export_as_xlsx.short_description = "Export Selected as XLSX"