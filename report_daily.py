import cx_Oracle
import pandas as pd
import win32com.client
from datetime import datetime


def run_daily_report():
    # Your database connection and data retrieval code here (same as before)
    dsn = cx_Oracle.makedsn("hostname", port_number, service_name="service_name")

    connection = cx_Oracle.connect(user='username', password='userpassword', dsn=dsn)
    query = """
    SELECT /* +parallel*/
        a.value_date,
        a.bank_date,
        c.pin,
        c.full_name,
        a.mcc,
        a.termlocation,
        a.currency AS CURRENCY_CODE,
        d.name AS currency_name,
        a.lcy_amount,
        a.fcy_amount,
        a.org_amount
    FROM
             cms.transaction_table a
        INNER JOIN shcema.customer_table b ON a.idclient = b.cms_id
        INNER JOIN shcema_name.employers_table  c ON b.pin = c.pin
        INNER JOIN shcema.currency_table d ON a.currency = d.id
    WHERE
        a.mcc IN ( '7800', '7801', '7802', '7995', '9754' )
    and a.value_date = TRUNC(SYSDATE)
    """
    df = pd.read_sql(query, con=connection)

    # Generate a file name with the current date
    current_date = datetime.now().strftime("%Y-%m-%d")  # Format as YYYY-MM-DD
    file_name = f"C:\\Users\\Username\\Desktop\\gambling\\gambling_everyday_{current_date}.xlsx"

    # Save DataFrame to the generated file
    df.to_excel(file_name, index=False)

    # Send the email with multiple CC and TO recipients
    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'Daily Report'

    # Multiple TO recipients (comma-separated)
    newmail.To = 'AnotherRecipient@example.com, YetAnotherRecipient@example.com'

    # Multiple CC recipients (comma-separated)
    newmail.CC = 'CCRecipient1@example.com, CCRecipient2@example.com'

    newmail.Body = 'Hello, this is a daily report email.'
    attach = file_name
    newmail.Attachments.Add(attach)
    newmail.Send()


if __name__ == "__main__":
    run_daily_report()
