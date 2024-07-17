import pandas as pd

#function to print responses
def print_company_responses(company_name, q1, q2, q3, q4, q5, q6, q7, q8, q9, q10):
    
    
    index = q1.index(company_name)
    
    print(f"Company name: {q1[index]}")

    if q2[index] == "Security":
        print("Q2: Microsoft Defender for Cloud helps you prevent, detect, and respond to threats with increased visibility into and control over the security of your Azure resources. Microsoft Sentinel delivers intelligent security analytics and threat intelligence. Azure Advisor is a personalized cloud consultant that helps you to optimize your Azure deployments. Learn more here: https://learn.microsoft.com/en-us/azure/security/fundamentals/overview#overview")
    elif q2[index] == "Data migration":
        print("Q2: Azure Database Migration Service is a fully managed service designed to enable seamless migrations from multiple ")
        print("database sources to Azure data platforms with minimal downtime. The Azure SQL Migration extension for Azure Data Studio brings together a simplified assessment, recommendation, and migration experience. Learn more here: https://learn.microsoft.com/en-us/azure/dms/dms-overview")
    elif q2[index] == "Storage":
        print("Q2: Azure provides secure and scallable storage solution. There are many different storage types to schoose fro which all have different benefits. Learn more here: https://learn.microsoft.com/en-us/azure/storage/common/storage-introduction")
    elif q2[index] == "Workplace modernisation":
        print("Q2: Microsoft 365 provides a wide array of applications to streamline your business. Applications such as Word, PowerPoint and excel can be used for processing documents and doing presentations. Teams and Outlook help provide messaging systems between your organisation. Learn more here: https://support.microsoft.com/en-us/office/microsoft-365-basics-video-training-396b8d9e-e118-42d0-8a0d-87d1f2f055fb")
    elif q2[index] == "App development":
        print("Q2: Avanade is experienced in making web applications and can develop one for you. To learn about the services we use to complete this visit: https://azure.microsoft.com/en-us/resources/developers ")
    else:
        print(f"Q2: Unknown response: {q2[index]}")

    if q3[index] == "agree":
        print("Q3: You will benefit from using a cloud database.")
    elif q3[index] == "disagree":
        print("Q3: A cloud database may help if you do experience any problems in the future.")
    elif q3[index] == "unsure":
        print("Q3: We can have a consultant look into data management for you")
    else:
        print(f"Q3: Unknown response: {q3[index]}")

    if q4[index] == "agree":
        print("Q4: You will benefit from using virtual machines or virtual storage")
    elif q4[index] == "disagree":
        print("Q4: Virtual machines or virtual storage may help if you do experience any problems in the future.")
    elif q4[index] == "unsure":
        print("Q4: We can have a consultant look into virtualisation for you")
    else:
        print(f"Q4: Unknown response: {q4[index]}")

    if q5[index] == "agree":
        print("Q5: We provide a wide array of security services which you can benefit from")
    elif q5[index] == "disagree":
        print("Q5: If you have any issues in the future, we provide a wide array of security services")
    elif q5[index] == "unsure":
        print("Q5: We can have a consultant look into security services for you")
    else:
        print(f"Q5: Unknown response: {q5[index]}")

    if q6[index] == "agree":
        print("Q6: You would benefit from using Azure sql database service")
    elif q6[index] == "disagree":
        print("Q6: Azure sql database service could provide solutions if you have any issues in the future")
    elif q6[index] == "unsure":
        print("Q6: We can have a consultant look into database options for you")
    else:
        print(f"Q6: Unknown response: {q6[index]}")

    if q7[index] == "agree":
        print("Q7: We can help you with any guidance you may need in the future")
    elif q7[index] == "disagree":
        print("Q7: Co-pilot is a useful tool that can be used across a wide array of Microsoft apps")
    elif q7[index] == "unsure":
        print("Q7: We can have a consultant look into AI for you")
    else:
        print(f"Q7: Unknown response: {q7[index]}")

    if q8[index] == "agree":
        print("Q8: Avanade is experienced in making web applications and can help you develop your own one")
    elif q8[index] == "disagree":
        print("Q8: If you need assistance with web apps in the future, we can help")
    elif q8[index] == "unsure":
        print("Q8: We can have a consultant look into applications for you")
    else:
        print(f"Q8: Unknown response: {q8[index]}")

    if q9[index] == "agree":
        print("Q9: We can help if you ever need guidance on your Microsoft systems")
    elif q9[index] == "disagree":
        print("Q9: Microsoft Teams and Outlook would be extremely useful for modernising your work communication methods")
    elif q9[index] == "unsure":
        print("Q9: We can have a consultant look into workplace modernisation for you")
    else:
        print(f"Q9: Unknown response: {q9[index]}")

    if q10[index] == "agree":
        print("Q10: It would be beneficial for you to be in contact with one of our experienced consultants")
    elif q10[index] == "disagree":
        print("Q10: You seem to be aware of the solutions you may need")
    elif q10[index] == "unsure":
        print("Q10: It would be beneficial for you to be in contact with one of our experienced consultants")
    else:
        print(f"Q10: Unknown response: {q10[index]}")
    
    print()
    

#excel file path
file_path = r"C:\Users\liam.cobb\Downloads\Untitled form(1-3).xlsx"

#read excel file
excel_data = pd.read_excel(file_path)

#print dataframe
print("DataFrame contents:")
print(excel_data)


q1 = []
q2 = []
q3 = []
q4 = []
q5 = []
q6 = []
q7 = []
q8 = []
q9 = []
q10 = []

#read through through rows
for index, row in excel_data.iterrows():
    #read through columns, start at 6
    for col_num, col_name in enumerate(excel_data.columns[6:], start=6):
        cell_value = row[col_name]

        #add values to list
        if col_num == 6:
            q1.append(cell_value)
        elif col_num == 7:
            q2.append(cell_value)
        elif col_num == 8:
            q3.append(cell_value)
        elif col_num == 9:
            q4.append(cell_value)
        elif col_num == 10:
            q5.append(cell_value)
        elif col_num == 11:
            q6.append(cell_value)
        elif col_num == 12:
            q7.append(cell_value)
        elif col_num == 13:
            q8.append(cell_value)
        elif col_num == 14:
            q9.append(cell_value)
        elif col_num == 15:
            q10.append(cell_value)

#display company names
print("List of company names:")
for i, company in enumerate(q1):
    print(f"{i + 1}. {company}")

#loop to catch errors
while True:
    try:
        company_number = int(input("Enter the number of your company: ")) - 1
        if 0 <= company_number < len(q1):
            company_name = q1[company_number]
            #print responses of company
            print_company_responses(company_name, q1, q2, q3, q4, q5, q6, q7, q8, q9, q10)
            break
        else:
            print("Invalid number. Please enter a valid number from the list.")
    except ValueError:
        print("Invalid input. Please enter a number.")
