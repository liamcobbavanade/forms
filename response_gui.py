import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import webbrowser

#function to print responses
def print_company_responses(company_name, q1, q2, q3, q4, q5, q6, q7, q8, q9, q10, q11):
    index = q1.index(company_name)
    
    response_text = ""
    
    response_text += f"Company name: {q1[index]}\n"
    response_text += f"Email: {q2[index]}\n\n"

    if q3[index] == "Security":
        response_text += ("Q3: Microsoft Defender for Cloud helps you prevent, detect, and respond to threats with increased visibility into and control over the security of your Azure resources. Microsoft Sentinel delivers intelligent security analytics and threat intelligence. Azure Advisor is a personalized cloud consultant that helps you to optimize your Azure deployments. Learn more here: https://learn.microsoft.com/en-us/azure/security/fundamentals/overview#overview\n\n")
    elif q3[index] == "Data migration":
        response_text += ("Q3: Azure Database Migration Service is a fully managed service designed to enable seamless migrations from multiple database sources to Azure data platforms with minimal downtime. The Azure SQL Migration extension for Azure Data Studio brings together a simplified assessment, recommendation, and migration experience. Learn more here: https://learn.microsoft.com/en-us/azure/dms/dms-overview\n\n")
    elif q3[index] == "Storage":
        response_text += ("Q3: Azure provides secure and scalable storage solutions. There are many different storage types to choose from, each with different benefits. Learn more here: https://learn.microsoft.com/en-us/azure/storage/common/storage-introduction\n\n")
    elif q3[index] == "Workplace modernisation":
        response_text += ("Q3: Microsoft 365 provides a wide array of applications to streamline your business. Applications such as Word, PowerPoint, and Excel can be used for processing documents and doing presentations. Teams and Outlook help provide messaging systems within your organization. Learn more here: https://support.microsoft.com/en-us/office/microsoft-365-basics-video-training-396b8d9e-e118-42d0-8a0d-87d1f2f055fb\n\n")
    elif q3[index] == "App development":
        response_text += ("Q3: Avanade is experienced in making web applications and can develop one for you. To learn about the services we use to complete this visit: https://azure.microsoft.com/en-us/resources/developers\n\n")
    else:
        response_text += (f"Q3: Unknown response: {q3[index]}\n\n")

    if q4[index] == "agree":
        response_text += ("Q4: Azure Database Migration Service is a fully managed service designed to enable seamless migrations from multiple database sources to Azure data platforms with minimal downtime. The Azure SQL Migration extension for Azure Data Studio brings together a simplified assessment, recommendation, and migration experience. Learn more here: https://azure.microsoft.com/en-us/products/database-migration/.\n\n")
    
    if q5[index] == "agree":
        response_text += ("Q5: A move to Cloud may benefit your business, remove costly hardware and replace with an easily manageable Cloud infrastructure. You can use virtual machines and virtual storage to deploy and save your data without the need for physical machines. https://azure.microsoft.com/en-gb/products/azure-migrate/.\n\n")
    
    if q6[index] == "agree":
        response_text += ("Q6: Microsoft Defender for Cloud helps you prevent, detect, and respond to threats with increased visibility into and control over the security of your Azure resources. Microsoft Sentinel delivers intelligent security analytics and threat intelligence. Azure Advisor is a personalized cloud consultant that helps you to optimize your Azure deployments. Learn more here: https://learn.microsoft.com/en-us/azure/security/fundamentals/overview#overview%22.\n\n")
    
    if q7[index] == "agree":
        response_text += ("Q7: Azure SQL Database is a managed cloud service by Microsoft Azure, ideal for modern applications. It features automatic scaling, AI assistance, mirroring, and robust security. Developers can leverage the latest SQL Server features without the overhead of managing infrastructure. https://azure.microsoft.com/en-us/products/azure-sql/database/.\n\n")
    
    if q8[index] == "agree":
        response_text += ("Q8: Integrating Microsoft Copilot within businesses enhances productivity and creativity. Key steps include activating Copilot for Microsoft 365, exploring its features, and ensuring priority access to the latest models. With careful planning, Copilot transforms work processes, freeing up time for growth and innovation. Learn more here: https://support.microsoft.com/en-gb/office/frequently-asked-questions-about-copilot-for-microsoft-365-500fc65e-9973-4e42-9cf4-bdefb0eb04ce.\n\n")
    
    if q9[index] == "agree":
        response_text += ("Q9: Avanade, a leading provider of digital, cloud, and advisory services, specializes in application development. They create custom applications that deliver engaging experiences, combining beautiful design with effective back-end integration. Avanade also offers managed services for maintaining existing applications and provides strategic advice in areas like sales, marketing, finance, and operations. https://www.avanade.com/en/solutions/cloud-and-application-services/development.\n\n")
    
    if q10[index] == "agree":
        response_text += ("Q10: Microsoft 365 is a cloud-powered productivity platform that includes apps like Word, Excel, PowerPoint, and Outlook. It’s available via subscription and offers cross-device installation, collaboration features, and cloud storage. Unlike Office 2021, it ensures up-to-date tools. Explore more at Microsoft 365. https://www.avanade.com/en/technologies/office-365.\n\n")
    
    if q11[index] == "agree":
        response_text += ("Q11: Speaking with a consultant will greatly benefit your company. Below is a link to Avanade’s website, navigate to “Contact us” and there you can fill out an inquiry form to send forward to a consultant. Link: https://www.avanade.com/en\n\n")
    
    return response_text

#function to handle button click event
def generate_responses():
    try:
        company_index = company_combo.current()
        if company_index >= 0:
            company_name = q1[company_index]
            response_text = print_company_responses(company_name, q1, q2, q3, q4, q5, q6, q7, q8, q9, q10, q11)
            response_textbox.config(state=tk.NORMAL)
            response_textbox.delete('1.0', tk.END)
            response_textbox.insert(tk.END, response_text)
            response_textbox.config(state=tk.DISABLED)
        else:
            messagebox.showwarning("Error", "Please select a company.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

#create GUI window
root = tk.Tk()
root.title("Company Responses Generator")

#read Excel file
file_path = r"C:\Users\liam.cobb\Downloads\email_form.xlsx"
excel_data = pd.read_excel(file_path)

#extract data from DataFrame
q1 = excel_data.iloc[:, 6].tolist()
q2 = excel_data.iloc[:, 7].tolist()
q3 = excel_data.iloc[:, 8].tolist()
q4 = excel_data.iloc[:, 9].tolist()
q5 = excel_data.iloc[:, 10].tolist()
q6 = excel_data.iloc[:, 11].tolist()
q7 = excel_data.iloc[:, 12].tolist()
q8 = excel_data.iloc[:, 13].tolist()
q9 = excel_data.iloc[:, 14].tolist()
q10 = excel_data.iloc[:, 15].tolist()
q11 = excel_data.iloc[:, 16].tolist()

#create widgets
ttk.Label(root, text="Select a Company:").pack(pady=10)
company_combo = ttk.Combobox(root, values=q1, state="readonly", width=50)
company_combo.pack()

ttk.Button(root, text="Generate Responses", command=generate_responses).pack(pady=10)

response_textbox = tk.Text(root, wrap=tk.WORD, width=100, height=20, state=tk.DISABLED)
response_textbox.pack(padx=10, pady=10)

#start GUI main loop
root.mainloop()
