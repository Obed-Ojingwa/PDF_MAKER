import utils.session
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Create a new Word Document
doc = Document()
style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(11)

# Add Contact Information
doc.add_heading('Obed Chukwunenye', level=1)
doc.add_paragraph("Rivers, Nigeria  | +234 810 254 4186  |  obedojingwa@gmail.com")
doc.add_paragraph("LinkedIn: https://www.linkedin.com/in/obed-ojingwa  |  GitHub: https://github.com/Obed-Ojingwa")

# Add Professional Summary
doc.add_heading("Professional Summary", level=2)
doc.add_paragraph(
    "Detail-oriented Data Analyst with proven skills in data wrangling, statistical analysis, dashboard creation, "
    "and data storytelling. Proficient in Python, SQL, Excel, Power BI, and Jupyter Notebooks. Experienced in "
    "transforming raw data into actionable insights that support strategic decision-making. Strong communicator with "
    "a passion for building scalable data solutions and visually intuitive dashboards. Actively building real-world "
    "projects and seeking to grow in a data-driven organization."
)

# Add Technical Skills
doc.add_heading("Technical Skills", level=2)
doc.add_paragraph(
    "• Languages: Python (Pandas, NumPy), SQL (SQLite, MySQL), VBA\n"
    "• BI & Analysis Tools: Excel (PivotTables, LOOKUPs), Google Sheets, Power BI, Tableau, \nMicrosoft BI, RapidMiner, KNIME, AWS QuickSight, Apache Spark\n"
    "• Visualization: Matplotlib, Seaborn, Plotly\n"
    "• Environments: Jupyter Notebook, Google Colab\n"
    "• Data Formats & APIs: CSV, JSON, XML, REST APIs\n"
    "• Version Control: Git & GitHub\n"
    "• Other: Web Scraping (BeautifulSoup, Requests), Markdown"
)

# Add Project Experience
doc.add_heading("Project Experience", level=2)

doc.add_paragraph("Retail Sales Analysis", style='List Bullet')
doc.add_paragraph("GitHub: https://github.com/Obed-Ojingwa/Retail_Sales_Analysis", style='Normal')
doc.add_paragraph(
    "Tools: Python, Jupyter Notebook, Excel, Power BI\n"
    "- Built a full-stack retail data pipeline: data collection, cleaning, exploratory analysis, and visualization.\n"
    "- Cleaned and structured transactional datasets using Pandas for accurate analysis.\n"
    "- Conducted sales trend analysis, customer segmentation, and revenue forecasting.\n"
    "- Delivered a dynamic Power BI dashboard to visualize top-performing products and monthly growth rates.\n"
    "- Achieved 25% accuracy boost in forecasting by applying moving averages and trendlines.\n"
    "- Communicated findings in a presentation-style Jupyter Notebook, highlighting actionable insights for stakeholders."
)

# Add Professional Experience
doc.add_heading("Professional Experience", level=2)

doc.add_paragraph("Data Analyst Intern (Freelance Projects) | Remote | Jan 2024 – July 2025", style='List Bullet')
doc.add_paragraph(
    "- Developed custom dashboards for small business clients using Excel and Google Sheets.\n"
    "- Automated monthly reporting using Python scripts and Excel macros, reducing manual effort by 60%.\n"
    "- Performed KPI tracking and data visualization for customer acquisition and retention analysis.\n"
    "- Created survey data summaries and presented statistical findings using bar charts, pie charts, and heatmaps."
)

doc.add_paragraph("GIS & Data Intern – IT | ASTAL | Akwa Ibom State, Nigeria | 2021 – 2022", style='List Bullet')
doc.add_paragraph(
    "- Supported spatial data collection and analysis using GIS, QGIS and Excel for urban planning purposes.\n"
    "- Designed thematic maps and geospatial reports used in environmental impact assessments.\n"
    "- Prepared tabular datasets and automated repetitive data entry tasks using VBA and formulas."
)

# Add Education
doc.add_heading("Education", level=2)
doc.add_paragraph("B.Sc. Geography and Planning | Abia State University, Nigeria | 2018 – 2022")
doc.add_paragraph("• Final Year Thesis: “Iddentification Of Groundwater Potential Zones In Abia State Using GIS Techniques”")
doc.add_paragraph("• Relevant Courses: Statistics, Remote Sensing, Spatial Analysis, Urban Modeling")

# Add Certifications
doc.add_heading("Certifications", level=2)
doc.add_paragraph(
    "• Data Analysis with Python – freeCodeCamp\n"
    "• Data Analysis With Python Full Course – Udemy \n"
    "• Excel for Business – Macquarie University (Coursera)\n"
    "• SQL for Data Analysts – Mode Analytics\n"
    "• Power BI Basics – YouTube Self-Learning (Dashboard Samples in GitHub)"
)

# Add Core Competencies
doc.add_heading("Core Competencies", level=2)
doc.add_paragraph(
    "Perseverance · Data Cleaning · Statistical Analysis · Data Visualization · Business Intelligence (BI) · KPI Reporting · "
    "Data Storytelling · Stakeholder Communication · Data Quality Assurance · Problem Solving · Time Management"
)

# Save the document
doc_path = "/home/nenye/Documents/PDF_MAKER/Obed_Chukwunenye_Data_Analyst_Resume.docx"
doc.save(doc_path)

doc_path
