# Планирование онда рабочего времени по загруженности косметологов
Скрипт анализирует соотношение планового и фактического (либо прогнозного для будущих дат) рабочее время косметологов, рассчитывает коэффициент загрузки сотрудников и формирует историю слепков расписания. Используется для нормирования труда, прогнозирования потребности в человеко-часах и планирования загрузки отделения косметологии.

# Cosmetologist Workload Analysis

> Python application for analyzing planned and actual cosmetologist workload and forecasting workforce capacity.

## Description

This project analyzes cosmetologist schedules by comparing planned working hours with actual completed appointments.

For future dates, the application estimates expected workload based on current bookings, allowing management to forecast staffing requirements and monitor workforce utilization.

The application also stores historical schedule snapshots, making it possible to compare workload changes over time.

Important:

The project requires access to the hospital database.

Without the original database structure the application cannot be executed.

## Business Goal

The application was developed as part of the workforce planning process.

It helps estimate how many standardized labor hours are scheduled for upcoming weeks and months, identify underloaded or overloaded employees and support staffing decisions.

## Features

- SQL database integration
- Planned working hours calculation
- Actual working hours calculation
- Forecast workload calculation
- Workforce utilization analysis
- Historical schedule snapshots
- Workload trend comparison
- Daily workload planning
- Excel report generation
- Conditional formatting

## Tech Stack

- Python
- pandas
- NumPy
- SQLAlchemy
- Firebird SQL
- xlsxwriter
- python-dotenv
- pyTelegramBotAPI

## How It Works

1. Loads employee schedules
2. Loads planned work shifts
3. Calculates actual worked time
4. Calculates planned working time
5. Computes workload percentage
6. Saves workload snapshot
7. Compares workload with previous snapshots
8. Generates Excel reports

## Example / Demo

### Input

Hospital database

- Employee schedules
- Planned shifts
- Appointment schedule

### Output

Excel workbook containing:

- Planned working hours
- Actual working hours
- Workload percentage
- Historical workload snapshots

## Use Case

This project can be used for:

- workforce planning
- workforce analytics
- healthcare analytics
- capacity planning
- operational reporting
