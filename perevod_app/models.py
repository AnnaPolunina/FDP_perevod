from django.db import models
from django.urls import reverse

class Translation(models.Model):
    en = models.CharField(
        max_length=50,
        unique=True, 
        verbose_name='Англійська'
        )
    uk = models.CharField(
        max_length=50,
        unique=True, 
        verbose_name='Українська'
        )
    
    def get_absolute_url(self):
        return reverse(self.detail_url_name, args=[self.id])

    def __str__(self):
        return '{0} {1} '.format(self.en, self.uk)
    
    class Meta:
        ordering = ['id']
        verbose_name = u'Переклад'
        verbose_name_plural = u'Переклади'

class ValueState(models.Model):
    excel_state = models.DecimalField(
        max_digits=1000000, 
        decimal_places=2, 
        )
    bd_state = models.DecimalField(
        max_digits=1000000, 
        decimal_places=2, 
        )