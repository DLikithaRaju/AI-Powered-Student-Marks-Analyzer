#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import speech_recognition as sr
import re
import pandas as pd
import os
from fpdf import FPDF
import matplotlib.pyplot as plt

# 📁 Ask the user to name the Excel file
def get_excel_filename():
    filename_input = input("📄 Enter the name for your Excel file (without extension): ").strip()
    if not filename_input.endswith(".xlsx"):
        filename_input += ".xlsx"
    filepath = os.path.join(os.path.expanduser("~"), "Desktop", filename_input)
    print(f"✅ Excel file will be saved at: {filepath}")
    return filepath

def clean_text(text):
    return re.sub(r'[^\x00-\x7F]+', '', text)

def get_voice_input():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("\n🎤 Speak now (e.g., Ravi got 45 in Math, 37 in Science, 89 in English)")
        try:
            audio = recognizer.listen(source, timeout=3, phrase_time_limit=6)
        except sr.WaitTimeoutError:
            print("❗ No speech detected in time.")
            return None

    try:
        text = recognizer.recognize_google(audio)
        print("✅ You said:", text)
        return text
    except Exception as e:
        print("❌ Speech recognition error:", e)
        return None

def parse_marks(text):
    result = {}
    name_match = re.search(r'(\w+)\s+(got|scored|has|obtained)', text, re.IGNORECASE)
    if name_match:
        result['Name'] = name_match.group(1)
    else:
        print("⚠️ Could not detect student name.")
        return None

    marks = re.findall(r'(\d+)\s+in\s+(\w+)', text, re.IGNORECASE)
    if not marks:
        marks = re.findall(r'(\d+)\s+(\w+)', text, re.IGNORECASE)
    if len(marks) == 0:
        print("⚠️ No valid subject-mark pairs found.")
        return None

    total = 0
    for mark, subject in marks:
        mark_int = int(mark)
        result[subject.capitalize()] = mark_int
        total += mark_int

    average = round(total / len(marks), 2)
    result['Total'] = total
    result['Average'] = average

    if average >= 90:
        result['Feedback'] = "Excellent performance! 🌟"
    elif average >= 75:
        result['Feedback'] = "Very good, keep it up! 👍"
    elif average >= 60:
        result['Feedback'] = "Good, but there's room to improve."
    elif average >= 40:
        result['Feedback'] = "Needs improvement. Focus more. 🔄"
    else:
        result['Feedback'] = "Poor performance. Seek help. 🆘"
    return result

def save_to_excel(data, filename):
    try:
        df_new = pd.DataFrame([data])
        if os.path.exists(filename):
            df = pd.read_excel(filename)
            df = pd.concat([df, df_new], ignore_index=True)
        else:
            df = df_new
        df.to_excel(filename, index=False)
        print(f"✅ Data saved to Excel at: {filename}")
    except Exception as e:
        print(f"❌ Error saving to Excel: {e}")

def generate_pdf_report(data, output_folder="Desktop"):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=14)

    pdf.cell(0, 10, f"Report Card - {data['Name']}", ln=True)
    for key, value in data.items():
        if key not in ["Name", "Total", "Average", "Feedback", "Rank"]:
            pdf.cell(0, 10, f"{key}: {value}", ln=True)

    pdf.cell(0, 10, f"Total: {data['Total']}", ln=True)
    pdf.cell(0, 10, f"Average: {data['Average']:.2f}", ln=True)
    pdf.cell(0, 10, f"Feedback: {clean_text(data['Feedback'])}", ln=True)
    if 'Rank' in data:
        pdf.cell(0, 10, f"Rank: {data['Rank']}", ln=True)

    output_path = os.path.join(os.path.expanduser("~"), output_folder, f"{data['Name']}_ReportCard.pdf")
    pdf.output(output_path)
    print(f"📄 PDF report saved at: {output_path}")

def collect_student_data(filename, student_limit=100):
    print(f"📚 Ready to enter marks for up to {student_limit} students.")
    print("🛑 Type 'stop' when asked if you want to continue.\n")

    for i in range(student_limit):
        print(f"\n🧑‍🎓 Student {i+1} - Speak now:")
        text = get_voice_input()

        if text:
            parsed_data = parse_marks(text)
            if parsed_data:
                print("📊 Parsed data:", parsed_data)
                save_to_excel(parsed_data, filename)
                generate_pdf_report(parsed_data)
                print("✅ Student data saved.")
            else:
                print("⚠️ Skipping due to invalid input.")
        else:
            print("⚠️ No input detected.")

        user_input = input("🔄 Do you want to continue? (yes/stop): ").strip().lower()
        if user_input == "stop":
            print("🛑 Stopping input as per user request.")
            break

def find_subject_wise_toppers(filename):
    try:
        df = pd.read_excel(filename)
        if df.empty:
            print("⚠️ No data found in Excel.")
            return

        print("\n🏆 Subject-wise Toppers:")
        subjects = [col for col in df.columns if col not in ["Name", "Total", "Average", "Feedback", "Rank"]]
        for subject in subjects:
            if pd.api.types.is_numeric_dtype(df[subject]):
                max_score = df[subject].max()
                toppers = df[df[subject] == max_score]['Name'].tolist()
                print(f"📚 {subject}: {', '.join(toppers)} (Score: {max_score})")
            else:
                print(f"⚠️ Skipping non-numeric subject: {subject}")
    except Exception as e:
        print(f"❌ Error finding toppers: {e}")

def generate_summary_report(filename, save_as_new=True):
    import pandas as pd
    import os

    df = pd.read_excel(filename)
    subjects = [col for col in df.columns if col not in ["Name", "Total", "Average", "Feedback", "Rank"]]

    # Calculate Total and Average
    df["Total"] = df[subjects].sum(axis=1)
    df["Average"] = df[subjects].mean(axis=1)

    # Calculate Rank
    df["Rank"] = df["Total"].rank(ascending=False, method='min').astype(int)

    # Generate Feedback based on average
    feedbacks = []
    for avg in df["Average"]:
        if avg >= 90:
            feedbacks.append("Excellent performance! 🌟")
        elif avg >= 75:
            feedbacks.append("Very good, keep it up! 👍")
        elif avg >= 60:
            feedbacks.append("Good, but there's room to improve.")
        elif avg >= 40:
            feedbacks.append("Needs improvement. Focus more. 🔄")
        else:
            feedbacks.append("Poor performance. Seek help. 🆘")
    df["Feedback"] = feedbacks

    # 🖨️ Print summary for each student
    print("\n📝 Detailed Feedback:")
    for _, row in df.iterrows():
        print(f"{row['Name']} - Total: {row['Total']}, Average: {row['Average']:.2f}, Rank: {row['Rank']}, Feedback: {row['Feedback']}")

    # 🏅 Subject-wise toppers
    print("\n📌 Subject-wise Toppers:")
    for subject in subjects:
        max_score = df[subject].max()
        toppers = df[df[subject] == max_score]["Name"].tolist()
        print(f"🏅 {subject}: {', '.join(toppers)}")

    # 🏆 Overall Rank List
    print("\n🏆 Overall Rank List:")
    print(df.sort_values("Rank")[["Name", "Total", "Average", "Rank"]].to_string(index=False))

    # 💾 Save enhanced Excel
    if save_as_new:
        new_filename = filename.replace(".xlsx", "_analyzed.xlsx")
        df.to_excel(new_filename, index=False)
        print(f"\n📁 Enhanced Excel saved as: {new_filename}")
    else:
        df.to_excel(filename, index=False)
        print(f"\n📁 Excel updated with new data at: {filename}")


import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd

# 🔍 Plot Top N Students by Total Marks

def plot_total_marks(df, top_n=10):
    top_df = df.sort_values("Total", ascending=False).head(top_n)

    plt.figure(figsize=(12, 6))
    bars = sns.barplot(x="Name", y="Total", data=top_df, palette="Blues_d")
    plt.title("Top Students by Total Marks", fontsize=16, weight='bold')
    plt.xlabel("Student Name")
    plt.ylabel("Total Marks")

    for bar in bars.patches:
        plt.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                 f'{bar.get_height():.0f}', ha='center', fontsize=10)

    plt.xticks(rotation=45)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.savefig("top_students_total_marks.png", dpi=300)
    plt.show()


# 📈 Subject-wise Average Line Chart

def plot_subject_averages(df):
    subjects = [col for col in df.columns if col not in ["Name", "Total", "Average", "Feedback", "Rank"]]
    averages = [df[subject].mean() for subject in subjects]

    plt.figure(figsize=(10, 5))
    plt.plot(subjects, averages, marker='o', linestyle='-', color='green')
    plt.title("Average Marks by Subject", fontsize=16, weight='bold')
    plt.xlabel("Subjects")
    plt.ylabel("Average Marks")
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.savefig("subject_average_line_chart.png", dpi=300)
    plt.show()


# 🥧 Feedback Distribution Pie Chart

def plot_feedback_distribution(df):
    feedback_counts = df["Feedback"].value_counts()

    plt.figure(figsize=(6, 6))
    plt.pie(feedback_counts, labels=feedback_counts.index, autopct="%1.1f%%",
            startangle=140, colors=sns.color_palette("pastel"))
    plt.title("Feedback Distribution", fontsize=14, weight='bold')
    plt.axis("equal")
    plt.tight_layout()
    plt.savefig("feedback_pie_chart.png", dpi=300)
    plt.show()


# 📊 Combined Visual Analysis Function

def generate_visual_charts(filename):
    df = pd.read_excel(filename)
    df.fillna(0, inplace=True)

    plot_total_marks(df)
    plot_subject_averages(df)
    plot_feedback_distribution(df)

    print("✅ Enhanced charts saved and displayed.")


def analyze_existing_excel():
    import time

    file_path = input("📄 Enter the full path to the existing Excel file: ").strip().strip('"').strip("'")
    if not os.path.exists(file_path):
        print("❌ File not found.")
        return

    print(f"\n📊 Analyzing: {file_path}")

    # ✅ First, enrich the Excel with Total, Average, Feedback, Rank
    generate_summary_report(file_path)

    # 🕒 Wait a tiny moment to ensure file save is complete (optional)
    time.sleep(0.5)

    # ✅ Read the updated Excel file
    df = pd.read_excel(file_path)

    # ✅ Check if 'Feedback' exists now
    if 'Feedback' not in df.columns:
        print("❌ 'Feedback' column still missing. Something went wrong in summary generation.")
        return

    # ✅ Generate PDF reports
    for _, row in df.iterrows():
        generate_pdf_report(row.to_dict())

    # ✅ Generate charts
    generate_visual_charts(file_path)

    print("\n📁 All report cards and analysis completed.")


    print("\n📁 All report cards and analysis completed.")

# -------------------------------
# 🚀 MAIN STARTS HERE
# -------------------------------

print("\n🎓 Welcome to Student Marks Analyzer")
print("1️⃣  Voice Input Mode (Live Entry)")
print("2️⃣  Analysis Mode (Existing Excel)")
choice = input("Select an option (1 or 2): ").strip()

if choice == "1":
    FILENAME = get_excel_filename()
    collect_student_data(FILENAME, student_limit=100)
    generate_summary_report(FILENAME,save_as_new=True)
    find_subject_wise_toppers(FILENAME)
    generate_visual_charts(FILENAME)
elif choice == "2":
    analyze_existing_excel()
else:
    print("❌ Invalid choice. Please select 1 or 2.")
    
    
# C:\Users\Owner\Desktop\analyzer_marks.xlsx

