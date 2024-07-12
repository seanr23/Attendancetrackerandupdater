from attendance import load_workbook, add_attendance, save_workbook
from email_utils import send_email


def main():
    file_path = 'attendance.xlsx'

    # Prompt user for input
    name = input('Enter the name: ')
    date = input('Enter the date (YYYY-MM-DD): ')
    status = input('Enter the status (Present/Absent): ')

    from_email = 'kkwc2ke@gmail.com'
    to_email = 'randolphseanj@gmail.com'
    email_password = 'rcmt iljf zisr uvrn'  # Use the generated app password here
    email_subject = 'Attendance Update'
    email_body = f'Attendance has been recorded for {name} on {date}: {status}'

    workbook, sheet = load_workbook(file_path)
    add_attendance(sheet, name, date, status)
    save_workbook(workbook, file_path)
    send_email(email_subject, email_body, to_email, from_email, email_password)

    print('Attendance recorded and email sent.')


if __name__ == '__main__':
    main()
