from flask import Flask, request, jsonify, render_template, send_from_directory
import pandas as pd
import numpy as np
import os
from openai import OpenAI
import json
from tqdm import tqdm

app = Flask(__name__)

# Connection to OpenAI API (requires OPENAI_API_KEY environment variable)
client = OpenAI()

@app.route('/')
def index():
    # Home page — file upload form
    return render_template('upload.html')


@app.route('/upload', methods=['POST'])
def upload_excel():
    # Receive two files — trees file and appellants file
    file1 = request.files.get('file1')
    file2 = request.files.get('file2')

    if not file1 or not file2:
        return jsonify({'error': 'שני קבצים נדרשים (עצים וערערים)'}), 400

    os.makedirs('uploads', exist_ok=True)
    os.makedirs('outputs', exist_ok=True)

    # ==========================
    # Step 1: Process trees file
    # ==========================
    path_trees = os.path.join('uploads', file1.filename)
    file1.save(path_trees)

    xl = pd.ExcelFile(path_trees)
    final_report = pd.DataFrame()

    # Iterate through all sheets (excluding auxiliary lists)
    for sheet_name in xl.sheet_names:
        if sheet_name not in ["רשימת עצים לפי קודים", "רשימת ערים לפי קודים"]:
            df = xl.parse(sheet_name, header=1)
            df["City!"] = sheet_name  # City name determined by sheet name

            # Standardize column names
            for name in ["כמות", "Quant", "כמות עצים", "סה'כ לכריתה", "מספר עצים"]:
                df.rename(columns={name: "numberOfTrees!"}, inplace=True)
            for name in ["סיבה", "Siba"]:
                df.rename(columns={name: "Siba!"}, inplace=True)
            for name in ["מין העץ", "Tree", "שם   מין עץ"]:
                df.rename(columns={name: "TreeName!"}, inplace=True)

            # Add default value if no reason column exists
            if "Siba!" not in df.columns:
                df["Siba!"] = "לא ידוע"

            # Keep only relevant columns
            df = df[["City!", "TreeName!", "Siba!", "numberOfTrees!"]].reset_index(drop=True)
            final_report = pd.concat([final_report, df], ignore_index=True)

    # Convert reason codes to readable text
    mapping = {
        1: "אחר", 2: "בטיחות", 3: "מחלת עץ", 4: "סכנה בריאותית", 5: "בנייה",
        6: "הכשרה חקלאית", 7: "עץ מת", 8: "דילול יער", 9: "קריאה", 10: "סניטציה", 11: "לא ידוע"
    }

    def convert_reason(val):
        try:
            num = int(float(str(val).strip()))
            return mapping.get(num, val)
        except (ValueError, TypeError):
            return val

    if "Siba!" in final_report.columns:
        final_report["Siba!"] = final_report["Siba!"].apply(convert_reason)
        final_report["Siba!"] = final_report["Siba!"].fillna("לא ידוע")
        final_report.loc[final_report["Siba!"].astype(str).str.strip() == "", "Siba!"] = "לא ידוע"

    # Merge with tree name list
    tree_codes = xl.parse("רשימת עצים לפי קודים", header=2)
    tree_codes.Tree = tree_codes.Tree.astype(str) + ".0"
    final_report["TreeName!"] = final_report["TreeName!"].astype(str)
    final_report = final_report.merge(tree_codes[["Tree", "שם עץ"]],
                                      left_on="TreeName!", right_on="Tree", how="left")
    final_report["TreeName!"] = np.where(final_report.Tree.notnull(),
                                         final_report["שם עץ"], final_report["TreeName!"])
    final_report.drop(columns=["Tree", "שם עץ"], inplace=True)

    # Clean column names
    final_report.columns = final_report.columns.str[:-1]
    output_trees = os.path.join('outputs', 'MergeFile.xlsx')
    final_report.to_excel(output_trees, index=False)

    # Dashboard calculations for trees file
    city_summary = final_report.groupby('City')['numberOfTrees'].sum().sort_values(ascending=False).head(10)
    top_cities = city_summary.reset_index().to_dict(orient='records')

    top_trees = (
        final_report.groupby('TreeName')['numberOfTrees'].sum()
        .sort_values(ascending=False).head(10)
        .reset_index().to_dict(orient='records')
    )

    top_reasons = (
        final_report['Siba'].astype(str).value_counts().head(10)
        .reset_index()
    )
    top_reasons.columns = ['Reason', 'Count']
    top_reasons = top_reasons.to_dict(orient='records')

    city_distribution = (
        final_report.groupby('City')['numberOfTrees'].sum()
        .sort_values(ascending=False)
        .reset_index().to_dict(orient='records')
    )

    # Top 10 cities with the highest number of cutting permits (each row = one permit)
    top_licenses = (
        final_report
        .groupby('City')
        .size()
        .sort_values(ascending=False)
        .head(10)
        .reset_index(name='LicenseCount')
        .to_dict(orient='records')
    )

    # Percent distribution of total cut trees by city (for pie chart)
    total_trees_all_cities = sum([c['numberOfTrees'] for c in city_distribution]) or 1 # avoid division by zero
    city_distribution_percent = [
        {
            'City': c['City'],
            'Percent': round((c['numberOfTrees'] / total_trees_all_cities) * 100, 1)
        }
        for c in city_distribution
    ]

    # =============================
    # Step 2: Process appellants file
    # =============================
    path_appellants = os.path.join('uploads', file2.filename)
    file2.save(path_appellants)

    xl_app = pd.ExcelFile(path_appellants)
    df_app = xl_app.parse(xl_app.sheet_names[0], skiprows=5)
    df_app = df_app.iloc[:, 2:6]
    df_app.columns = ["כתובת", "סיבת הבקשה", "החלטת פקיד אזורי", "החלטת פקיד ממשלתי"]

    # Add additional columns for processing
    df_app["ישוב"] = ""
    df_app["מספר עצים שנכרתו"] = 0
    df_app["מספר עצים שנשמרו"] = 0
    df_app["G"] = ""
    df_app["H"] = ""

    print("מתחיל ניתוח GPT על קובץ הערערים...")

    max_rows = 221

    for i, row in tqdm(df_app.iterrows(), total=min(len(df_app), max_rows)):
        if i >= max_rows:
            print(f"הגעתי לשורה {i+1} — עוצר, אין צורך לעבור מעבר ל-221 רשומות.")
            break

        address = str(row["כתובת"]).strip()
        local_decision = str(row["החלטת פקיד אזורי"]).strip()
        appeal_decision = str(row["החלטת פקיד ממשלתי"]).strip()

        if not address and not local_decision and not appeal_decision:
            print(f"עצירה בשורה {i+1}: אין נתונים נוספים — סוף הרשומות.")
            break

        # GPT prompt — analyze the row and infer conclusions
        prompt = f"""
        אתה עוזר אנליסט לקריאת החלטות כריתת עצים בישראל.
        הנתונים לשורה:
        כתובת: "{address}"
        החלטת פקיד אזורי: "{local_decision}"
        החלטת פקיד ממשלתי (ערעור): "{appeal_decision}"

        בצע את הפעולות הבאות:
        1. מצא את שם העיר מתוך הכתובת בלבד.
        2. קבע G:
           "Y" אם הפקיד האזורי החליט לכרות עצים.
           "N" אם הפקיד האזורי החליט שלא לכרות.
        3. קבע H:
           - "YY" אם G="Y" והערעור התקבל
           - "NN" אם G="N" והערעור התקבל
           - "YN" אם G="Y" והערעור נדחה
           - "NY" אם G="N" והערעור נדחה
        4. חשב כמה עצים נכרתו בפועל וכמה נשמרו.
        החזר אך ורק JSON בפורמט הבא:
        {{
            "city": "<שם העיר>",
            "G": "<Y/N>",
            "H": "<YY/NN/YN/NY>",
            "cut": <int>,
            "saved": <int>
        }}
        """

        try:
            response = client.responses.create(
                model="gpt-4.1-mini",
                input=prompt,
                temperature=0
            )
            text = response.output_text.strip().strip("`")
            if text.startswith("json"):
                text = text[4:].strip()
            result = json.loads(text)

            df_app.at[i, "ישוב"] = result.get("city", "")
            df_app.at[i, "G"] = result.get("G", "")
            df_app.at[i, "H"] = result.get("H", "")
            df_app.at[i, "מספר עצים שנכרתו"] = result.get("cut", 0)
            df_app.at[i, "מספר עצים שנשמרו"] = result.get("saved", 0)

        except Exception as e:
            print(f"שגיאה בשורה {i+1}: {e}")
            df_app.at[i, "מספר עצים שנכרתו"] = 0
            df_app.at[i, "מספר עצים שנשמרו"] = 0

    # Save the processed appellants file
    print("ניתוח GPT הושלם, שומר קובץ ערערים...")
    output_appellants = os.path.join('outputs', 'Appellants_Analyzed.xlsx')
    df_app.to_excel(output_appellants, index=False)

    # =============================
    # Additional analysis for appeals
    # =============================
    try:
        appeals_df = df_app.copy()

        # Top 10 cities with most appeals (G="Y")
        appeal_cities = (
            appeals_df[appeals_df["G"] == "Y"]
            .groupby("ישוב").size()
            .reset_index(name="count")
            .sort_values(by="count", ascending=False)
            .head(10)
            .rename(columns={"ישוב": "city"})
            .to_dict(orient="records")
        )

        # Top 10 successful appeals (H="YY")
        top_successful_appeals = (
            appeals_df[appeals_df["H"] == "YY"]
            .sort_values(by="מספר עצים שנשמרו", ascending=False)
            .head(10)[["ישוב", "מספר עצים שנשמרו"]]
            .rename(columns={"ישוב": "city", "מספר עצים שנשמרו": "saved"})
            .to_dict(orient="records")
        )

        # Common reasons for successful appeals (H="YY")
        successful_appeals = appeals_df[appeals_df["H"] == "YY"]
        if not successful_appeals.empty:
            reason_counts = (
                successful_appeals["סיבת הבקשה"]
                .astype(str)
                .value_counts()
                .head(10)
                .reset_index()
            )
            reason_counts.columns = ["reason", "count"]
            appeal_reasons = reason_counts.to_dict(orient="records")
        else:
            appeal_reasons = []

    except Exception as e:
        print(f"שגיאה בעת ניתוח עררים: {e}")
        appeal_cities, top_successful_appeals, appeal_reasons = [], [], []

    print("כל הנתונים מוכנים — מחזיר תגובה ללקוח.")

    # Return JSON response to frontend
    return jsonify({
        'message': 'שני הקבצים עובדו והוערכו בהצלחה',
        'trees_file': output_trees,
        'appellants_file': output_appellants,
        'top_cities': top_cities,
        'top_trees': top_trees,
        'top_reasons': top_reasons,
        'city_distribution': city_distribution,
        'top_licenses': top_licenses, 
        'appeal_cities': appeal_cities,
        'top_successful_appeals': top_successful_appeals,
        'appeal_reasons': appeal_reasons,
        'city_distribution_percent': city_distribution_percent
    })


@app.route('/download/<path:filename>')
def download_file(filename):
    # Download files from 'outputs' directory
    return send_from_directory('outputs', filename, as_attachment=True)


if __name__ == '__main__':
    # Create folders if not existing and start server
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('outputs', exist_ok=True)
    app.run(debug=True)
