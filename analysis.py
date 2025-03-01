import pandas as pd

data = [
    ["ARUN KUMAWAT", "0827CA21DD12", "Integrated MCA", "CA", 7, "Regular", "C+", "A", "A"],
    ["AYUSH UPADHYAY", "0827CA21DD16", "Integrated MCA", "CA", 7, "Regular", "C", "B+", "B+"],
    ["CHETAN MAKWANA", "0827CA21DD18", "Integrated MCA", "CA", 7, "Regular", "C+", "A", "A"]
]

columns = ["Name", "Roll No.", "Course", "Branch", "Semester", "Status", "MCADD705- [T]", "MCADD706- [P]", "MCADD707- [P]"]

# Convert list of lists into DataFrame
df = pd.DataFrame(data, columns=columns)

# Extract only subject columns
subjects = df.columns[6:]  # Selecting subject columns only

# Define required grades order
grades = ["A+", "A", "B+", "B", "C+", "C", "D", "F"]

# Create grade distribution dictionary
grade_distribution = {subject: df[subject].value_counts() for subject in subjects}

# Convert to DataFrame and fill missing grades with 0
grade_df = pd.DataFrame(grade_distribution).fillna(0).astype(int)

# Reorder rows to match required grade order
grade_df = grade_df.reindex(grades, fill_value=0)

grade_avg = grade_df.mean()

grade_df.loc["Average"] = grade_avg.round(2)  # Average count of each grade across subjects

print(grade_df)
