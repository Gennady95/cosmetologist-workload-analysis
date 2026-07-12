import pandas as pd, re, os, glob, xlsxwriter, telebot, getpass, platform, time, threading, requests, numpy as np, locale
from datetime import datetime, timedelta, date
from sqlalchemy import create_engine
from dotenv import load_dotenv

# Паттерны
load_dotenv()
pd.set_option('display.max_columns', None)
pd.set_option('mode.chained_assignment', None)
pd.options.mode.chained_assignment = None
# Создание коннекторов данных
start_time = time.time() # Текущее время
datename = datetime.now().strftime('%d.%m %H:%M') # Время создания файла
bot = telebot.TeleBot(os.getenv("TELEGRAM_TOKEN"))
engine = create_engine(os.getenv("DB_URL"))
pattern = "Слепок загруженности косметологов.xlsx"                                                                                                                                                       # Паттерн для поиска уже сформированного файла

def SendTelegram(status, er):
    UserName = getpass.getuser()                                                                                                                                                                         # Имя пользователя (обычно оно User - не информативно)
    CompName = platform.node()                                                                                                                                                                           # Имя компьютера
    chat_id = '5249664773'                                                                                                                                                                               # ID моей телеги
    if status == "try": # Если связь с телегой установлена
        bot.send_message(chat_id, datename+" пользователь "+UserName+" ("+CompName+") успешно воспользовался скриптом для подсчёта клиентов с болезнями")                                                # Отправка сообщения
    elif status == "except1": # Если нет подключения к SQL серверу
        bot.send_message(chat_id, "ERROR: " + datename+" пользователь "+UserName+" ("+CompName+") неудачно запустил скрипт для подсчёта клиентов с болезнями: " + er)                                                  # Отправка сообщения
def convert_colname(col):
	if isinstance(col, (datetime, date)):
		return col.strftime('%d.%m.%y')
	if isinstance(col, str):
		match = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', col)
		if match:
			return f"{match.group(3)}.{match.group(2)}.{match.group(1)[2:]}"
	return col
def process_df_for_concat(new_df, old_df):
	# Все колонки в строку, чтобы не крашилось
	new_df.columns = new_df.columns.astype(str)
	old_df.columns = old_df.columns.astype(str)
	
	non_date_cols = [col for col in new_df.columns if not re.match(r'\d{2}\.\d{2}\.\d{2}$', col)]
	date_cols = sorted(
		[col for col in new_df.columns if re.match(r'\d{2}\.\d{2}\.\d{2}$', col)],
		key=lambda x: datetime.strptime(x, "%d.%m.%y")
	)
	ordered_cols = non_date_cols + date_cols
	
	new_df = new_df.reindex(columns=ordered_cols)
	old_df = old_df.reindex(columns=ordered_cols)
	
	# Создаём пустую строку с пустыми строками в Имя доктора и Дата слепка, а в остальных nan
	empty_row = pd.DataFrame([{col: "" if col in ['Имя доктора', 'Дата слепка'] else np.nan for col in ordered_cols}])
	
	merged_df = pd.concat([new_df, empty_row, old_df], ignore_index=True)
	
	# Заполняем nan в ключевых колонках пустыми строками
	for col in ['Имя доктора', 'Дата слепка']:
		if col in merged_df.columns:
			merged_df[col] = merged_df[col].fillna('')
	
	return merged_df
def add_summary_row(df, aggfunc, name="Имя доктора"):
	if df.empty: return df
	summary = {}
	for col in df.columns:
		if col == name: summary[col] = "Итог по отделению"
		elif col == 'Дата слепка': summary[col] = ""
		else: summary[col] = aggfunc(pd.to_numeric(df[col], errors='coerce'))
	return pd.concat([df, pd.DataFrame([summary])], ignore_index=True)
def add_weekdays_row(df):
	days_map = {'mon': 'пн', 'tue': 'вт', 'wed': 'ср', 'thu': 'чт', 'fri': 'пт', 'sat': 'сб', 'sun': 'вс'}
	date_cols = [c for c in df.columns if re.match(r"\d{2}\.\d{2}\.\d{2}", str(c))]
	
	# Формируем строку с днями недели, в ключевых колонках - пусто
	weekday_row = []
	for c in df.columns:
		if c in date_cols:
			day_eng = datetime.strptime(c, "%d.%m.%y").strftime("%a").lower()
			weekday_row.append(days_map.get(day_eng, ''))
		else:
			weekday_row.append('')  # пустая строка для нефактовых колонок
	
	# Создаем DataFrame из этой строки
	weekday_df = pd.DataFrame([weekday_row], columns=df.columns)
	
	# Принудительно приводим типы ключевых столбцов к str и вставляем пустые строки в этой строке
	for col in ['Имя доктора', 'Дата слепка']:
		if col in weekday_df.columns:
			weekday_df[col] = weekday_df[col].astype(str).replace('nan', '')
	
	# Объединяем
	df = pd.concat([weekday_df, df], ignore_index=True)
	
	# ВАЖНО: после конкатенации столбцы могут быть числовыми, а в нулевой строке у нас пустые строки - нужно привести к str
	for col in ['Имя доктора', 'Дата слепка']:
		if col in df.columns:
			# Заменяем nan и делаем тип строкой (для всей колонки, чтобы избежать nan)
			df[col] = df[col].astype(str).replace('nan', '')
	
	return df
def write_change_sheet(file_path):
	import pandas as pd, numpy as np

	# читаем последний и предпоследний лист
	sheets = pd.ExcelFile(file_path).sheet_names
	if len(sheets) < 2:
		print("Недостаточно листов для сравнения изменений.")
		return

	prev_df = pd.read_excel(file_path, sheet_name=sheets[-2]).set_index('ФИО')
	curr_df = pd.read_excel(file_path, sheet_name=sheets[-1]).set_index('ФИО')

	# только числовые столбцы (датированные)
	prev = prev_df.select_dtypes(include='number')
	curr = curr_df.select_dtypes(include='number')

	# объединяем все даты-столбцы и врачи
	all_cols = sorted(set(prev.columns).union(curr.columns), key=lambda d: pd.to_datetime(d, dayfirst=True))
	all_doctors = sorted(set(prev.index).union(curr.index))

	prev = prev.reindex(index=all_doctors, columns=all_cols, fill_value=np.nan)
	curr = curr.reindex(index=all_doctors, columns=all_cols, fill_value=np.nan)

	# считаем разницу
	diff = curr - prev
	diff = diff.round(1).astype(str) + '%'
	diff[curr.isna() | prev.isna()] = ''  # чистим там, где не было данных

	diff.reset_index(inplace=True)

	with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
		diff.to_excel(writer, sheet_name='Изменения', index=False)
def GetSQL():
	global doctor_list, start_date, end_date, start_time, old_data_plan, old_data_fact, old_data_workload

	# SQL запросы
	q_schedule = "SELECT schedid, dcode, dcode1, workdate, bhour, bmin, fhour, fmin, depnum, pcode, chid, status, clvisit, COMMENT from SCHEDULE"
	q_doctors = "SELECT dcode, dname, depnum, lockdate from BI_DOCTORS WHERE lockdate IS NULL"
	q_doctshed = "SELECT dcode, wdate, beghour, begmin, endhour, endmin from DOCTSHEDULE WHERE depnum = '40001404'"
	SCHEDULE = pd.read_sql(q_schedule, engine);	BI_DOCTORS = pd.read_sql(q_doctors, engine); DOCTSHEDULE = pd.read_sql(q_doctshed, engine)
	print(f"Все базы загружены за: {(time.time() - start_time):.2f}"); start_time = time.time()
	
	# Форматирование массива
	doctor_list = BI_DOCTORS[BI_DOCTORS['depnum'] == 40001404]['dname'].dropna().tolist()
	name_map = dict(BI_DOCTORS[['dcode', 'dname']].values)
	SCHEDULE['dcode'] = SCHEDULE['dcode'].replace(-1, np.nan)
	SCHEDULE['Доктор'] = SCHEDULE[['dcode', 'dcode1']].bfill(axis=1).iloc[:, 0]
	SCHEDULE['Имя доктора'] = SCHEDULE['Доктор'].map(name_map)
	DOCTSHEDULE['Имя доктора'] = DOCTSHEDULE['dcode'].map(name_map)
	SCHEDULE = SCHEDULE[SCHEDULE['Имя доктора'].isin(doctor_list)]
	SCHEDULE['Продолжительность работы'] = (SCHEDULE['fhour'] - SCHEDULE['bhour']) * 60 + (SCHEDULE['fmin'] - SCHEDULE['bmin'])
	DOCTSHEDULE['Продолжительность смены'] = (DOCTSHEDULE['endhour'] - DOCTSHEDULE['beghour']) * 60 + (DOCTSHEDULE['endmin'] - DOCTSHEDULE['begmin'])

	# Выбор диапазона дат
	start_date = datetime.now().replace(hour=0, minute=0)
	end_date = DOCTSHEDULE['wdate'].max().replace(hour=0, minute=0)
	print(f"Дата начала: {start_date}\nДата окончания: {end_date}")
	filtered_schedule = SCHEDULE[(SCHEDULE['workdate'] >= start_date) & (SCHEDULE['workdate'] <= end_date)].copy()
	filtered_doctshed = DOCTSHEDULE[(DOCTSHEDULE['wdate'] >= start_date) & (DOCTSHEDULE['wdate'] <= end_date)].copy()
	all_dates_range = pd.date_range(start=start_date, end=end_date, freq='D') # Создаём полный список дат от start_date до end_date (включительно)
	all_dates_str = [d.strftime('%d.%m.%y') for d in all_dates_range]
	# Группировка по датам
	plan_raw = filtered_doctshed.groupby(['Имя доктора', filtered_doctshed['wdate'].dt.strftime('%d.%m.%y')])['Продолжительность смены'].sum().unstack(fill_value=0)
	fact_raw = filtered_schedule.groupby(['Имя доктора', filtered_schedule['workdate'].dt.strftime('%d.%m.%y')])['Продолжительность работы'].sum().unstack(fill_value=0)
	# Приведение к полному списку врачей и дат
	all_doctors = sorted(set(plan_raw.index) | set(fact_raw.index))
	plan = plan_raw.reindex(index=all_doctors, columns=all_dates_str, fill_value=0).reset_index()
	fact = fact_raw.reindex(index=all_doctors, columns=all_dates_str, fill_value=0).reset_index()
	all_doctors = sorted(set(plan['Имя доктора']) | set(fact['Имя доктора']))
	all_dates = sorted(set(plan.columns[1:]) | set(fact.columns[1:]))

	# Основная часть работы
	workload_rows = []
	for doc in all_doctors:
		row = {'Имя доктора': doc}
		p = plan[plan['Имя доктора'] == doc]
		f = fact[fact['Имя доктора'] == doc]
		for d in all_dates:
			val_p = p[d].iloc[0] if d in p.columns and not p.empty else 0
			val_f = f[d].iloc[0] if d in f.columns and not f.empty else 0
			row[d] = val_f / val_p if val_p else 0
		row['Дата слепка'] = datename
		workload_rows.append(row)
	workload = pd.DataFrame(workload_rows)
	if not workload.empty:
		for d in all_dates:
			if d not in workload.columns: workload[d] = np.nan
		workload = workload[['Имя доктора', 'Дата слепка'] + all_dates]
	else: workload = pd.DataFrame(columns=['Имя доктора', 'Дата слепка'])
	
	plan = add_summary_row(plan, lambda x: x.sum(skipna=True))
	fact = add_summary_row(fact, lambda x: x.sum(skipna=True))
	workload = add_summary_row(workload, lambda x: x.mean(skipna=True))
	plan.insert(1, 'Дата слепка', datename)
	fact.insert(1, 'Дата слепка', datename)
	final_plan_df = process_df_for_concat(plan, old_data_plan)
	final_fact_df = process_df_for_concat(fact, old_data_fact)
	date_cols = [col for col in final_plan_df.columns[2:] if col in final_fact_df.columns] # Берём даты, которые есть и там, и там
	
	for col in date_cols:
		final_plan_df[col] = pd.to_numeric(final_plan_df[col], errors='coerce')
		final_fact_df[col] = pd.to_numeric(final_fact_df[col], errors='coerce')
	
	final_workload_df = final_plan_df.copy()
	for col in date_cols:
		final_workload_df[col] = np.where(final_plan_df[col].fillna(0) > 0, final_fact_df[col].fillna(0) / final_plan_df[col], np.nan)
	
	# Добавление строки с днями недели
	final_plan_df = add_weekdays_row(final_plan_df)
	final_fact_df = add_weekdays_row(final_fact_df)
	final_workload_df = add_weekdays_row(final_workload_df)
	
	# Запись в файл
	with pd.ExcelWriter(pattern, engine='xlsxwriter') as writer:
		workbook = writer.book
		# Загруженность
		final_workload_df.rename(columns=convert_colname, inplace=True)
		final_workload_df.to_excel(writer, sheet_name='Загруженность', index=False, freeze_panes=(2, 2))
		worksheet = writer.sheets['Загруженность']
		percent_fmt = writer.book.add_format({'num_format': '0.0%', 'align': 'center'})
		if not final_workload_df.empty:
			worksheet.set_column(0, 0, 22); worksheet.set_column(1, 1, 10)
			worksheet.autofilter(0, 0, final_workload_df.shape[0], final_workload_df.shape[1] - 1)
			worksheet.set_column(2, final_workload_df.shape[1] - 1, 7, percent_fmt)
			worksheet.conditional_format(1, 2, final_workload_df.shape[0], final_workload_df.shape[1] - 1, {'type': '3_color_scale', 'min_type': 'num', 'min_value': 0.01, 'min_color': '#8B0000', 'mid_type': 'num', 'mid_value': 0.5, 'mid_color': '#FFFF66', 'max_type': 'num', 'max_value': 0.9, 'max_color': '#90EE90',})
			nrows, ncols = final_plan_df.shape[0] + 1, final_plan_df.shape[1]
			for r in range(nrows): worksheet.set_row(r, 25 if r == 0 or final_plan_df.iloc[r - 1].isnull().all() else 15) # Для строки заголовков и пустых разделительных строк высота = 20, для остальных = 10
		# Рабочие часы
		final_fact_df.rename(columns=convert_colname, inplace=True)
		final_fact_df.to_excel(writer, sheet_name='Рабочие часы', index=False, freeze_panes=(2, 2))
		worksheet = writer.sheets['Рабочие часы']
		if not final_fact_df.empty:
			worksheet.set_column(0, 0, 22); worksheet.set_column(1, 1, 10)
			worksheet.set_column(2, final_fact_df.shape[1] - 1, 6)
			worksheet.autofilter(0, 0, final_fact_df.shape[0], final_fact_df.shape[1] - 1)
			nrows, ncols = final_plan_df.shape[0] + 1, final_plan_df.shape[1]
			for r in range(nrows): worksheet.set_row(r, 25 if r == 0 or final_plan_df.iloc[r - 1].isnull().all() else 15)  # Для строки заголовков и пустых разделительных строк высота = 20, для остальных = 10
		# Плановые часы
		final_plan_df.rename(columns=convert_colname, inplace=True)
		final_plan_df.to_excel(writer, sheet_name='Плановые часы', index=False, freeze_panes=(2, 2))
		worksheet = writer.sheets['Плановые часы']
		if not final_plan_df.empty:
			worksheet.set_column(0, 0, 22); worksheet.set_column(1, 1, 10)
			worksheet.set_column(2, final_plan_df.shape[1] - 1, 6)
			worksheet.autofilter(0, 0, final_plan_df.shape[0], final_plan_df.shape[1] - 1)
			nrows, ncols = final_plan_df.shape[0] + 1, final_plan_df.shape[1]
			for r in range(nrows): worksheet.set_row(r, 25 if r == 0 or final_plan_df.iloc[r - 1].isnull().all() else 15)  # Для строки заголовков и пустых разделительных строк высота = 20, для остальных = 10
	final_workload_df = add_weekdays_row(final_workload_df)
# Запуск скрипта, проверка наличия файла для старта
try:
	files = glob.glob(pattern)
	last_file = max(files, key=os.path.getmtime)
	print(f"Найден файл: {last_file}")
	xlsx = pd.ExcelFile(last_file)
	old_data_plan = xlsx.parse('Плановые часы')
	old_data_fact = xlsx.parse('Рабочие часы')
	old_data_workload = xlsx.parse('Загруженность')

	for df in [old_data_plan, old_data_fact, old_data_workload]:
		for col in ['Имя доктора', 'Дата слепка']:
			if col in df.columns:
				df[col] = df[col].astype(str)
		new_col_names = {}
		for col in df.columns:
			if isinstance(col, (datetime, date)):
				new_col_names[col] = col.strftime('%d.%m.%y')
			elif isinstance(col, str):
				match = re.match(r'^(\d{4})-(\d{2})-(\d{2})', col)
				if match:
					new_col_names[col] = f"{match.group(3)}.{match.group(2)}.{match.group(1)[2:]}"
		df.rename(columns=new_col_names, inplace=True)
	print("Скрипт работает по алгоритму пополнения")
except:
	old_data_plan = pd.DataFrame(); old_data_fact = pd.DataFrame(); old_data_workload = pd.DataFrame()
	print("Файл не найден, старт с нуля")

GetSQL()