from django.contrib import admin
from .models import Translation

class TranslationAdmin(admin.ModelAdmin):
    
    list_display = (
        'en',
        'uk'
        )
    
    search_fields = (
        'en',
        'uk'
        )
    
admin.site.register(Translation, TranslationAdmin)
