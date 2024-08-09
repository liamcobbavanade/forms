import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#print responses
def print_company_responses(content_array):
    content_array = [item.strip() for item in content_array]

    response_text = ""
    response_text += f"Hello {content_array[0]},\n\n"
    response_text += "Thank you for filling out our quick quote service form. Here are some resources to help you understand the solutions you may need.\n\n"

    if content_array[2] == "Security":
        response_text += ("Q3: Microsoft Defender for Cloud helps you prevent, detect, and respond to threats with increased visibility into and control over the security of your Azure resources. Microsoft Sentinel delivers intelligent security analytics and threat intelligence. Azure Advisor is a personalized cloud consultant that helps you to optimize your Azure deployments. Learn more here: https://learn.microsoft.com/en-us/azure/security/fundamentals/overview#overview\n\n")
    elif content_array[2] == "Data migration":
        response_text += ("Q3: Azure Database Migration Service is a fully managed service designed to enable seamless migrations from multiple database sources to Azure data platforms with minimal downtime. The Azure SQL Migration extension for Azure Data Studio brings together a simplified assessment, recommendation, and migration experience. Learn more here: https://learn.microsoft.com/en-us/azure/dms/dms-overview\n\n")
    elif content_array[2] == "Storage":
        response_text += ("Q3: Azure provides secure and scalable storage solutions. There are many different storage types to choose from, each with different benefits. Learn more here: https://learn.microsoft.com/en-us/azure/storage/common/storage-introduction\n\n")
    elif content_array[2] == "Workplace modernisation":
        response_text += ("Q3: Microsoft 365 provides a wide array of applications to streamline your business. Applications such as Word, PowerPoint, and Excel can be used for processing documents and doing presentations. Teams and Outlook help provide messaging systems within your organization. Learn more here: https://support.microsoft.com/en-us/office/microsoft-365-basics-video-training-396b8d9e-e118-42d0-8a0d-87d1f2f055fb\n\n")
    elif content_array[2] == "App development":
        response_text += ("Q3: Avanade is experienced in making web applications and can develop one for you. To learn about the services we use to complete this visit: https://azure.microsoft.com/en-us/resources/developers\n\n")
    else:
        response_text += (f"Q3: Unknown response: {content_array[3]}\n\n")

    if content_array[3] == "agree":
        response_text += ("Q4: Azure Database Migration Service is a fully managed service designed to enable seamless migrations from multiple database sources to Azure data platforms with minimal downtime. The Azure SQL Migration extension for Azure Data Studio brings together a simplified assessment, recommendation, and migration experience. Learn more here: https://azure.microsoft.com/en-us/products/database-migration/.\n\n")

    if content_array[4] == "agree":
        response_text += ("Q5: A move to Cloud may benefit your business, remove costly hardware and replace with an easily manageable Cloud infrastructure. You can use virtual machines and virtual storage to deploy and save your data without the need for physical machines. https://azure.microsoft.com/en-gb/products/azure-migrate/.\n\n")

    if content_array[5] == "agree":
        response_text += ("Q6: Microsoft Defender for Cloud helps you prevent, detect, and respond to threats with increased visibility into and control over the security of your Azure resources. Microsoft Sentinel delivers intelligent security analytics and threat intelligence. Azure Advisor is a personalized cloud consultant that helps you to optimize your Azure deployments. Learn more here: https://learn.microsoft.com/en-us/azure/security/fundamentals/overview#overview%22.\n\n")

    if content_array[6] == "agree":
        response_text += ("Q7: Azure SQL Database is a managed cloud service by Microsoft Azure, ideal for modern applications. It features automatic scaling, AI assistance, mirroring, and robust security. Developers can leverage the latest SQL Server features without the overhead of managing infrastructure. https://azure.microsoft.com/en-us/products/azure-sql/database/.\n\n")

    if content_array[7] == "agree":
        response_text += ("Q8: Integrating Microsoft Copilot within businesses enhances productivity and creativity. Key steps include activating Copilot for Microsoft 365, exploring its features, and ensuring priority access to the latest models. With careful planning, Copilot transforms work processes, freeing up time for growth and innovation. Learn more here: https://support.microsoft.com/en-gb/office/frequently-asked-questions-about-copilot-for-microsoft-365-500fc65e-9973-4e42-9cf4-bdefb0eb04ce.\n\n")

    if content_array[8] == "agree":
        response_text += ("Q9: Avanade, a leading provider of digital, cloud, and advisory services, specializes in application development. They create custom applications that deliver engaging experiences, combining beautiful design with effective back-end integration. Avanade also offers managed services for maintaining existing applications and provides strategic advice in areas like sales, marketing, finance, and operations. https://www.avanade.com/en/solutions/cloud-and-application-services/development.\n\n")

    if content_array[9] == "agree":
        response_text += ("Q10: Microsoft 365 is a cloud-powered productivity platform that includes apps like Word, Excel, PowerPoint, and Outlook. It’s available via subscription and offers cross-device installation, collaboration features, and cloud storage. Unlike Office 2021, it ensures up-to-date tools. Explore more at Microsoft 365. https://www.avanade.com/en/technologies/office-365.\n\n")

    if content_array[10] == "agree":
        response_text += ("Q11: Speaking with a consultant will greatly benefit your company. Below is a link to Avanade’s website, navigate to “Contact us” and there you can fill out an inquiry form to send forward to a consultant. Link: https://www.avanade.com/en\n\n")


    response_text += "\nThank you for filling out your inquiry, we hope to hear from you again soon.\n\n"
    response_text += "Kind regards,\n"
    response_text += "The Avanade team\n"

    return response_text

#send email
def send_email(receiver_email, subject, body):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    sender_email = ''
    password = '' 

    #headers
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    #add body to email
    message.attach(MIMEText(body, 'plain'))

    try:
        #connect to server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  

        #log in to server
        server.login(sender_email, password)

        #send email
        server.sendmail(sender_email, receiver_email, message.as_string())

        print(f'Email sent successfully to {receiver_email}!')

    except Exception as e:
        print(f'Error: {e}')

    finally:
        server.quit()

spark.conf.set(
    "fs.azure.account.key.storageavanadeinterns.dfs.core.windows.net", "<secret-key>")

dbutils.fs.ls("abfss://container@storageavanadeinterns.dfs.core.windows.net/clientresponses/")

#read file into dataframe
df = spark.read.text("abfss://container@storageavanadeinterns.dfs.core.windows.net/clientresponses/clientform")

#get top row into string
full_text = "\n".join(row.value for row in df.collect())

#split string by comma
content_array = full_text.split(",")

#print array
print(content_array)

response_text = print_company_responses(content_array)
receiver_email = content_array[1]
send_email(receiver_email, f'Responses for {content_array[0]}', response_text)
