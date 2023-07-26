from datetime import datetime, timedelta
import pytz

def check_schedule():
    
    # 日本の時刻
    timezone_utc_plus_9 = pytz.timezone('Asia/Tokyo')
    currrent_time = timezone_utc_plus_9 = pytz.timezone('Asia/Tokyo')
    
# 毎週金曜日のスケジュール
    friday_schedule = {'day_of_week': 4, 'start_time': '03:00', 'end_time': '03:30'}
    
    # 毎月第2水曜日のスケジュール
    second_wednesday = current_time.replace(day=15)
    while second_wednesday.weekday() != 2:  # 水曜日を探す
        second_wednesday += timedelta(days=1)
        second_wednesday_schedule = {'day': second_wednesday.day, 'start_time': '22:30', 'end_time': '23:59'}
    next_day_schedule = {'day': (second_wednesday+timedelta(days=1)).day, 'start_time': '00:00', 'end_time': '08:00'}
    
    # 現在時刻がスケジュール内かどうかをチェック
    schedules = [friday_schedule, second_wednesday_schedule, next_day_schedule]
    
    for schedule in schedules:
        start_time = datetime.strptime(schedule['start_time'], '%H:%M').time()
        end_time = datetime.strptime(schedule['end_time'], '%H:%M').time()

        if not (day := schedule.get("day")):
            # 毎週金曜日について
            day_of_week = schedule['day_of_week']
        
            if current_time.weekday() == day_of_week and start_time <= current_time.time() <= end_time:
                return 0
        else:
            #第2水曜日の判定　
            if current_time.day == day and start_time <= current_time.time() <= end_time:
              return 0
    return 1
