# 🌳 Trees Dashboard - Final Project

An interactive data analysis and visualization system for forestry data in Israel. The dashboard automatically processes and analyzes Excel reports of tree-cutting permits and public appeals, extracts structured insights using OpenAI’s GPT-4.1-mini, and generates clear visual summaries through interactive charts.

This project was developed as part of the final B.Sc. Computer Science project at Ono Academic College. It demonstrates the integration of AI reasoning, data cleaning, and visualization to support transparent decision-making for environmental management.

- - -

## ⚙️ How to Run

1. Install dependencies:
pip install flask pandas numpy tqdm openai python-dotenv

2. Set your OpenAI API key:  
For macOS / Linux:  
export OPENAI_API_KEY="your_api_key_here"  
For Windows PowerShell:  
setx OPENAI_API_KEY "your_api_key_here"  

4. Run the Flask server:
python main.py

5. Open your browser at:
http://127.0.0.1:5000

Then upload two Excel files - one for tree-cutting permits and one for appeals.  
The dashboard will generate insights and graphs automatically.

You can easily download the files here (2024):  
https://www.gov.il/he/pages/ararim_2023
https://www.gov.il/he/pages/tree_clearing_license_2024

---

## 🧰 Technologies

Backend: Flask, Pandas, NumPy  
AI Integration: OpenAI GPT-4.1-mini  
Frontend: HTML, CSS, JS, Chart.js  
Visualization: Bar & Pie charts  
Execution: Local Flask server

---

## 👩‍💻 Authors

Gil Hadad & Inbar Abraham  
Ono Academic College, Department of Computer Science, 2025
