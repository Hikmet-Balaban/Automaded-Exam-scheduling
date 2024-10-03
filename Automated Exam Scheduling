import pandas as pd
import datetime

# 1. Önceden Belirlenmiş Sınav Programını Yükleme
predefined_schedule_path = r'C:\Users\hikme\Desktop\Bitirme Projesi\Dekanlık Sınav programı (1).xlsx'
predefined_schedule = pd.read_excel(predefined_schedule_path)

# Sütun adlarını kontrol edelim
print(predefined_schedule.columns)

# 2. Zaman Dilimlerini Dinamik Olarak Tanımlama
dates = predefined_schedule['SINAV TARİHİ'].unique()
schedule = {}

for date in dates:
    exam_date_str = pd.to_datetime(date).strftime('%Y-%m-%d')
    time_slots = predefined_schedule[predefined_schedule['SINAV TARİHİ'] == date]['SINAV SAATİ'].unique()
    schedule[exam_date_str] = {slot: [] for slot in time_slots}  

# 3. Önceden Belirlenmiş Sınavları Yerleştirme
predefined_courses = set()

for index, row in predefined_schedule.iterrows():
    exam_date = row['SINAV TARİHİ'] 
    time_slot = row['SINAV SAATİ']  
    course = row['DERS KODU']  
    
    exam_date_str = pd.to_datetime(exam_date).strftime('%Y-%m-%d')
    
    if exam_date_str in schedule and time_slot in schedule[exam_date_str]:
        schedule[exam_date_str][time_slot].append(course) 
        predefined_courses.add(course)

# schedule yapısını kontrol edelim
print(f"Schedule: {schedule}")

# 4. Çakışma Matrisini Yükleme
conflict_matrix_path = r'C:\Users\hikme\Desktop\Bitirme Projesi\Kitle Sınavları - Çakışma Matrisi Güz 2023.xlsx'
conflict_matrix = pd.read_excel(conflict_matrix_path, index_col=0)

# 5. Ders Kodu Prefix Tespiti (Dinamik Prefix Kontrolü)
def get_course_prefix(course_code):
    return ''.join([char for char in course_code if not char.isdigit()])

# 6. Tüm Zaman Dilimlerini Yaratmak
def generate_full_schedule_with_all_time_slots(schedule):
    all_time_slots = set()
    for date in schedule:
        all_time_slots.update(schedule[date].keys())
    
    all_time_slots = sorted(list(all_time_slots))  

    for date in schedule:
        for time_slot in all_time_slots:
            if time_slot not in schedule[date]:
                schedule[date][time_slot] = []  
    
    return schedule

# 7. Çakışma Matrisine Göre Dersleri Yerleştirme Fonksiyonu
def place_courses_with_no_conflict(schedule, conflict_matrix, max_occurrences_per_day=3):
    for course in conflict_matrix.index:
        if pd.isna(course) or not isinstance(course, str):
            continue  
        
        course_prefix = get_course_prefix(course)  
        
        placed = False
        for date in schedule.keys():
            course_count = sum([sum([1 for c in slot if get_course_prefix(c) == course_prefix]) for slot in schedule[date].values()])
            if course_count >= max_occurrences_per_day:
                continue  
                
            for time_slot in schedule[date].keys():
                conflicts = False
                for scheduled_course in schedule[date][time_slot]:
                    if scheduled_course in conflict_matrix.columns:
                        conflict_value = conflict_matrix.at[course, scheduled_course]
                        if conflict_value > 0:
                            conflicts = True
                            break
                
                if not conflicts:  
                    schedule[date][time_slot].append(course)
                    print(f"Yerleştirilen Ders: {course}, Tarih: {date}, Zaman Dilimi: {time_slot}")
                    placed = True
                    break
            if placed:
                break

# 8. Tüm zaman dilimlerini ekleyelim
schedule = generate_full_schedule_with_all_time_slots(schedule)

# 9. Mevcut programda yerleştirilmiş dersleri kaydediyoruz
predefined_courses = set([course for date in schedule for slot in schedule[date] for course in schedule[date][slot]])

# Çakışma matrisine göre dersleri yerleştirme
place_courses_with_no_conflict(schedule, conflict_matrix)

# 10. Yeni Ders Programını Kaydetme
output_file_path = r'C:\Users\hikme\Desktop\Bitirme Projesi\Yeni_Ders_Programı.xlsx'

output_data = []
for date, slots in schedule.items():
    for time_slot, courses in slots.items():
        if courses:
            for course in courses:
                output_data.append([date, time_slot, course])

# DataFrame oluşturma ve Excel'e kaydetme
if output_data:  
    output_df = pd.DataFrame(output_data, columns=['Tarih', 'Zaman Dilimi', 'Ders Kodu'])
    output_df.to_excel(output_file_path, index=False)
    print(f'Ders programı başarıyla {output_file_path} dosyasına kaydedildi.')
else:
    print("Output Data boş, Excel dosyası oluşturulmadı.")

# 11. Gerçekçi Bir Ders Programı Tablosu Oluşturma
time_slots = set()
for date in schedule:
    time_slots.update(schedule[date].keys())

time_slots = sorted(list(time_slots))  
schedule_table = pd.DataFrame(index=time_slots, columns=schedule.keys())

for date in schedule.keys():
    for time_slot in schedule[date].keys():
        courses_at_time = schedule[date][time_slot]
        
        if courses_at_time:
            schedule_table.at[time_slot, date] = ', '.join(courses_at_time)

# 12. Excel Dosyasına Kaydetme
output_file_path = r'C:\Users\hikme\Desktop\Bitirme Projesi\Final_Ders_Programı.xlsx'
schedule_table.to_excel(output_file_path)

print(f'Ders programı başarıyla {output_file_path} dosyasına kaydedildi.')
