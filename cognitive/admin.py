from django.contrib import admin
from .models import User, TestNSI, TestResult


@admin.register(User)
class UserAdmin(admin.ModelAdmin):
    list_display = ('username', 'age', 'education', 'speciality')


@admin.register(TestNSI)
class TestNSIAdmin(admin.ModelAdmin):
    list_display = ('test_name', 'title_all', 'title_correct')


@admin.register(TestResult)
class TestResultAdmin(admin.ModelAdmin):
    list_display = ('user', 'test', 'try_number', 'accuracy')
    actions = ['export_as_csv', 'export_as_xlsx']

    def export_as_csv(self, request, queryset):
        import csv
        from django.http import HttpResponse
        response = HttpResponse(content_type='text/csv')
        response['Content-Disposition'] = 'attachment; filename="test_results.csv"'
        writer = csv.writer(response)
        writer.writerow(['User', 'Test', 'Try Number', 'Accuracy'])
        for result in queryset:
            writer.writerow([result.user.username, result.test.test_name, result.try_number, result.accuracy])
        return response

    def export_as_xlsx(self, request, queryset):
        import openpyxl
        from django.http import HttpResponse
        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename="test_results.xlsx"'
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['User', 'Test', 'Try Number', 'Accuracy'])
        for result in queryset:
            ws.append([result.user.username, result.test.test_name, result.try_number, result.accuracy])
        wb.save(response)
        return response

    export_as_csv.short_description = "Export Selected as CSV"
    export_as_xlsx.short_description = "Export Selected as XLSX"
