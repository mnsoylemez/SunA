import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import pvlib 
from pvlib import location
from pvlib import irradiance
import datetime
import os 

def calculate_and_export_energy(locations_data, start_year, panel_efficiency_decimal, num_years=10, filepath=None):

    if not filepath:
        messagebox.showerror("Hata", "Kayıt dosya yolu belirtilmedi.")
        return
    if not (0 < panel_efficiency_decimal <= 1.0):
         messagebox.showerror("Girdi Hatası", "Panel verimliliği %0 ile %100 arasında olmalıdır.")
         return

    try:
        start_date = pd.Timestamp(f'{start_year}-01-01 00:00:00', tz='UTC')
        end_date = pd.Timestamp(f'{start_year + num_years}-01-01 00:00:00', tz='UTC')
        # Use inclusive='left' for pandas >= 2.2.0
        times_utc = pd.date_range(start_date, end_date, freq='h', inclusive='left')

        all_energy_data = {} 

        for loc_data in locations_data:
            print(f"{loc_data['name']} işleniyor...")
            try:
                loc = location.Location(
                    latitude=loc_data['latitude'],
                    longitude=loc_data['longitude'],
                    tz=loc_data['timezone'], 
                    altitude=0,
                    name=loc_data['name']
                )
                
                times_local = times_utc.tz_convert(loc_data['timezone']) 
                solpos = loc.get_solarposition(times_local)
                clearsky = loc.get_clearsky(times_local, model='ineichen', solar_position=solpos)

               
                generated_power_w = clearsky['ghi'] * panel_efficiency_decimal

                
                hourly_energy_wh = generated_power_w
                hourly_energy_wh.name = f"{loc_data['name']} Enerji (Wh)" 
                all_energy_data[loc_data['name']] = hourly_energy_wh

            except Exception as e:
                messagebox.showerror("Hesaplama Hatası", f"{loc_data['name']} konumu işlenirken hata oluştu:\n{e}")
                return

        
        if not all_energy_data:
            messagebox.showinfo("Veri Yok", "Enerji verisi hesaplanamadı.")
            return

        combined_df = pd.DataFrame(all_energy_data)

        
        monthly_total_wh = combined_df.resample('M').sum()

        
        monthly_total_kwh = monthly_total_wh / 1000.0

        
        monthly_total_kwh.columns = [f"{name} Açık Hava Enerji (kWh/ay, %{panel_efficiency_decimal*100:.1f} verim)" for name in monthly_total_kwh.columns]

        
        monthly_total_kwh.index = monthly_total_kwh.index.strftime('%Y-%m')

        
        print(f"Aylık enerji üretimi şuraya aktarılıyor: {filepath}")
        sheet_name = f'Aylik_kWh_{panel_efficiency_decimal*100:.0f}verim'
        if len(sheet_name) > 31: 
            sheet_name = sheet_name[:31]
        monthly_total_kwh.to_excel(filepath, sheet_name=sheet_name)
        print("Dışa aktarma tamamlandı.")
        messagebox.showinfo("Başarılı", f"Aylık açık gökyüzünde enerji üretim verileri şuraya kaydedildi:\n{filepath}\nSayfa: {sheet_name}\n\nNot: Bu, ideal açık gökyüzünde koşullarında 1m² panel için potansiyel enerjiyi temsil eder.")

    except ValueError as ve:
         messagebox.showerror("Girdi Hatası", f"Lütfen girdi değerlerini kontrol edin.\nDetaylar: {ve}")
    except Exception as e:
        messagebox.showerror("Hata", f"Beklenmedik bir hata oluştu:\n{e}")
        import traceback
        traceback.print_exc()


root = tk.Tk()
root.title("Güneş Paneli Enerji Üretimi Karşılaştırması (1m²)")


input_frame = ttk.LabelFrame(root, text="Girdiler", padding=(10, 5))
input_frame.pack(padx=10, pady=10, fill="x")



tz_display_options = []
tz_value_map = {}
for offset in range(-12, 15):
    display = f"GMT{offset:+d}"

    value = f"Etc/GMT{-offset:+d}"
    tz_display_options.append(display)
    tz_value_map[display] = value

default_tz_display = "GMT+3" 


locations_entries = []
def add_location_fields(frame, loc_num):
    
    loc_frame = ttk.Frame(frame)
    loc_frame.pack(fill="x", pady=5)
    ttk.Label(loc_frame, text=f"Konum {loc_num}:").grid(row=0, column=0, columnspan=4, sticky="w", padx=5)

    ttk.Label(loc_frame, text="İsim:").grid(row=1, column=0, sticky="w", padx=5)
    name_entry = ttk.Entry(loc_frame, width=15)
    name_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=5) 

    ttk.Label(loc_frame, text="Enlem (°):").grid(row=2, column=0, sticky="w", padx=5)
    lat_entry = ttk.Entry(loc_frame, width=10)
    lat_entry.grid(row=2, column=1, sticky="ew", padx=5)

    ttk.Label(loc_frame, text="Boylam (°):").grid(row=2, column=2, sticky="w", padx=5)
    lon_entry = ttk.Entry(loc_frame, width=10)
    lon_entry.grid(row=2, column=3, sticky="ew", padx=5)

    ttk.Label(loc_frame, text="Saat Dilimi:").grid(row=3, column=0, sticky="w", padx=5)
    
    tz_combo = ttk.Combobox(loc_frame, values=tz_display_options, state="readonly", width=10)
    tz_combo.grid(row=3, column=1, columnspan=3, sticky="ew", padx=5)
    tz_combo.set(default_tz_display)

    loc_frame.columnconfigure(1, weight=1)
    loc_frame.columnconfigure(3, weight=1) 

    return {'name': name_entry, 'lat': lat_entry, 'lon': lon_entry, 'tz_combo': tz_combo} 

locations_entries.append(add_location_fields(input_frame, 1))
locations_entries.append(add_location_fields(input_frame, 2))


config_frame = ttk.Frame(input_frame)
config_frame.pack(fill="x", pady=5, padx=5)

ttk.Label(config_frame, text="Başlangıç Yılı:").pack(side=tk.LEFT, padx=5)
year_entry = ttk.Entry(config_frame, width=8)
year_entry.pack(side=tk.LEFT)
current_year = datetime.datetime.now().year
year_entry.insert(0, str(current_year - 1))

ttk.Label(config_frame, text="Panel Verimliliği (%):").pack(side=tk.LEFT, padx=(20, 5))
efficiency_entry = ttk.Entry(config_frame, width=6)
efficiency_entry.pack(side=tk.LEFT)
efficiency_entry.insert(0, "20") 

ttk.Label(config_frame, text="(10 yıl için)").pack(side=tk.LEFT, padx=10)



def on_calculate_click():
    
    locations_data = []
    try:
        start_year = int(year_entry.get())
        efficiency_percent = float(efficiency_entry.get())

        if not (0 < efficiency_percent <= 100): 
             raise ValueError("Panel Verimliliği %0 ile %100 arasında olmalıdır.")
        panel_efficiency_decimal = efficiency_percent / 100.0

        if start_year < 1950 or start_year > 2100:
            raise ValueError("Yıl 1950 ile 2100 arasında olmalıdır")

        for i, entries in enumerate(locations_entries):
            name = entries['name'].get().strip()
            lat_str = entries['lat'].get()
            lon_str = entries['lon'].get()
            selected_display_tz = entries['tz_combo'].get() 

            if not all([name, lat_str, lon_str, selected_display_tz]):
                 messagebox.showerror("Girdi Hatası", f"Lütfen Konum {i+1} için tüm alanları doldurun.")
                 return

            
            try:
                 tz_value = tz_value_map[selected_display_tz]
            except KeyError:
            
                 messagebox.showerror("Girdi Hatası", f"Konum {i+1} için geçersiz saat dilimi seçildi.")
                 return

            locations_data.append({
                'name': name,
                'latitude': float(lat_str),
                'longitude': float(lon_str),
                'timezone': tz_value
            })

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel dosyaları", "*.xlsx"), ("Tüm dosyalar", "*.*")],
            title="Aylık Enerji (kWh) Verisini Farklı Kaydet..."
        )

        if save_path:
            calculate_and_export_energy(
                locations_data,
                start_year,
                panel_efficiency_decimal,
                num_years=10,
                filepath=save_path
            )
        else:
            print("Kaydetme işlemi iptal edildi.")

    except ValueError as ve:
        messagebox.showerror("Girdi Hatası", f"Lütfen Yıl, Verimlilik, Enlem ve Boylam için geçerli sayılar girin.\nDetaylar: {ve}")
    except Exception as e:
        messagebox.showerror("Hata", f"Girdi işlenirken beklenmedik bir hata oluştu:\n{e}")
        import traceback
        traceback.print_exc()


calculate_button = ttk.Button(root, text="Hesapla ve Aylık Enerjiyi (kWh) Dışa Aktar", command=on_calculate_click)
calculate_button.pack(padx=10, pady=10)


help_label = ttk.Label(root, text="1m² yatay panel için potansiyel açık gökyüzünde enerji üretimini (kWh/ay) hesaplar.\nSaat dilimi yukarıdaki listeden seçilmelidir.", wraplength=450, justify=tk.CENTER) # Translated help text
help_label.pack(padx=10, pady=5)


root.mainloop()