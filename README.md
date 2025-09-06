# TechnoEdge Power BI Automation Project – Email Attachments to Automated Dashboards to stakeholders

## 📌 Project Overview
In this **Power BI automation project** i will explain how businesses can save time and resources by automaticlly ingesting the data in power BI, making reports and sending to the stakeholders.

This project is related to **FinTech sector**, but this automation in power bi is applicable in other organizations as well. In many organizations teams receive data in the form of email attachments (Excel, CSV, PDF, JSON, etc.) daily or multiple times a day. In normal practice, a team of analysts work in this way that they:
- Download each attachment manually  
- clean and combine the files  
- Create reports in Excel  
- Share the results with stakeholders  

So it takes a huge amount of time and resources. There also chances of errors.  

👉 In this Project, I have made an automation for this whole process that will use **Microsoft Power BI** and **Outlook integration**.  
👉 In this automation an email with an excel attachment received in outlook, it will automatically comes in Power BI, automatically cleaned and transformed, updating the visuals in report and this updated report will automatically sent to the stakeholders in their email inbox.

👉 Porcess: "email received in outlook" ---> "comes in Power BI" ---> "cleaned & transformed" ---> "report updated with new data "---> "report send to stakholder"

---

## 🚀 Problem Statement
- Data updates frequently through email attachments (it may be hourly or monthly reports).  
- For analysis, "Annually months of efforts" is need to perform manual repetitive tasks. 
- Delay in report, chances of errors, and conficting.  

---

## ✅ Solution
In Microsoft Power BI get data through Microsoft Exchange Online:
1. **Automatically Retrieve Email Attachments**  
   - Outlook/Exchange can be connected directly to the Power BI.  
   - There's no need to download or open attachments by hand.  

2. **Transform & Clean Data**  
   - Combine several files (CSV, Excel and others).  
   - Manages schema consistency (requred same column names).  
   - Data that is not required is excluded (for example, 2019 data is ignored).  

3. **Creating automated Dashboard**  
   - Make very beautiful visualizations in Power BI using the combined dataset.  
   - When new attachment comes in outlook, the report is automatically refreshed and updated.  

4. **Delivery to the stakeholders Automatically**  
   - The Report is shareable by using **Power BI Service**.  
   - Also we can dilver the report to the stakeholders, automatically, at a scheduled day and time with a message and other settings, using the "Add Subscriptions" option in Power BI.  

---

## 🔧 Requirements
- The most recent Power BI Desktop version is advised but you can aslo you any previous reliable version.  
- A Corporate email is required for Microsoft Exchange Online, Support for Gmail or Yahoo is not available. It will attach the corporate email with Power BI.
- Understanding the basics of how to filter and arrange emails in folders using Outlook "Rules" option.

---

## 📂 Our Data
Monthly files in our case contained the following data:
- **Dates** → Ship Date, Order Date
- **Text fields** → Ship Mode, Segment, Customer Name
- **Geographical data** → Region, City, Postal Code, Country
- **Category fields** → Product Name, Category, Sub-Category   
- **Numeric fields** → Profit, Quantity, Sales, Discount

---

## ⚙️ Steps involved
1. **Create Outlook Rule** → Emails that are related to our requirements on the basis of subject will be moved into a specfic seperate folder(make a new one).  
2. **Line Power BI to Exchange Online** → Import attachments and mail data.  
3. **Filter Attachments** → Ignore all the old data which is not required (for example 2019) and keep only CSV or Excle files that are needed.  
4. **Extract Data** → Use a function to extract data, in our case it is "Excel.Workbook()". For other file types this function is little different like for csv it is "CSV.Document(), for pdf it would be PDF.Tables().
5. **Combine Datasets** → All files will be appended into a single dataset to make a single Combined dataset.  
6. **Create Dashboard** → Make beautiful Visuals for profitabiltiy, trends, sales and regional insights.
7. **Third party Visual** → I have also added a third party visual called scroller for more insightfull report creation. This is a visual you have mostly seen in news channels and stock markets.
8. **Publish to Power BI Service** → Add subscriptions for the stakeholders to whom you want to sent report at scheduled day and time. so the stakeholders can automatically receive the report through email daily or weekly. Also set scheduled refresh, whenever a new email comes in outlook with the specific subject line, the dataset will be refreshed automatically and the report will be updated.

---

## 🎯 Business Impact
- Saves **4–5 months of work per analyst annually**.  
- Minimize human error and manual efforts.  
- Offers stakeholders automated report in real-time.  
- Easily expands to thousands of email attachments.

---

## 📸 Report
![Report](Resources/atuomated%20sales%20pbi%20screenshot.png)

---

## Link of Project on Power BI Services



## 📚 Learnings

- Making a spereate folder and shifting the files meeting a specific criteria to that new folder in outlook, using the "Rule" option in "Home Tab".
- How to extract different kinds of data from email in Power BI.
- Extraction of Excel content in Power BI from email.
- Adding 3rd Party visuals in Dashboards like scroller.
- Giving presentation using the full screen option in Power BI. 
- How to automate end-to-end reporting pipelines in Power BI. 
- Connecting the Power BI with Outlook/Exchange Online.  
- Handling email-based semi-structured data ingestion.  
- Delivery of automated report to clients or managers directly in their inbox using the "Add subscription" option in Power BI.

---

## 🛠️ Tech Stack
- **Power BI Desktop & Service**  
- **Microsoft Exchange Online (Outlook)**  
- **CSV / Excel data sources**  

---

## 📌 Author
Developed by **Zaghem Abbas** – Data Analyst | Business Intelligence Enthusiast  

---

