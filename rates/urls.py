from django.urls import path
from . import views

urlpatterns = [
    path('', views.success_page, name='home'),
    path('rates/', views.show_rates, name='rates'),
    path('import/', views.import_excel, name='import_excel'),
    path('scraper/', views.scraper_page, name='scraper'),
    path('run-scraper/', views.run_scraper, name='run_scraper'),
    path('get-ports/', views.get_ports, name='get_ports'),
    path('download-excel/', views.download_excel, name='download_excel'),
    path('run-tracker/', views.tracker_page, name='tracker_page'),
    path('run-tracker/run/', views.run_tracker, name='run_tracker'),
    path('get-shipping-lines/', views.get_shipping_lines, name='get_shipping_lines'),
    path('stop-selected-scrapers/', views.stop_selected_scrapers, name='stop_selected_scrapers'),
    path('stop-all/', views.stop_all_scrapers, name='stop_all_scrapers'),
    path('get-running-scrapers/', views.get_running_scrapers, name='get_running_scrapers'),
]