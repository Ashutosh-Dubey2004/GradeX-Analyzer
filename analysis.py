import pandas as pd

def get_grade_from_average(avg):
    """Returns the grade equivalent for the given average score."""
    grade_scale = [(9.0, "A+"), (8.0, "A"), (7.0, "B+"), (6.0, "B"), (5.0, "C+"), (4.0, "C"), (3.0, "D")]
    for threshold, grade in grade_scale:
        if avg >= threshold:
            return grade
    return "F"

def perform_analysis(df):
    if df.empty:
        print("Dataset is empty.")
        return None

    df.columns = df.columns.str.strip()  # Ensure column names are stripped of spaces
    subject_columns = df.columns[6:-3]  # Extract subjects ]
    if subject_columns.empty:
        print("No subject columns detected. Please check dataset format.")
        return None

    df[subject_columns] = df[subject_columns].fillna("F").astype(str)

    grade_weights = {"A+": 10, "A": 9, "B+": 8, "B": 7, "C+": 6, "C": 5, "D": 4, "F": 0}
    
    # Grade Distribution Table
    # grade_distribution = {subject: df[subject].value_counts() for subject in subject_columns}
    # grade_df = pd.DataFrame(grade_distribution).fillna(0).astype(int)  

    # # Ensure all grade rows are present in correct order
    # grade_df = grade_df.reindex(grade_weights.keys(), fill_value=0).astype(int)
    # total_students = grade_df.sum(axis=0).astype(int)  
    # grade_df.loc["Total Appeared"] = total_students

    grade_df = df[subject_columns].apply(pd.Series.value_counts).fillna(0).astype(int)
    grade_df = grade_df.reindex(grade_weights.keys(), fill_value=0)
    grade_df.loc["Total Appeared"] = grade_df.sum()
    
    # Weighted Average Grade Calculation
    # weighted_sums = sum(grade_df.loc[grade] * weight for grade, weight in grade_weights.items())
    # weighted_avg = (weighted_sums / total_students).fillna(0).round(2)

    weighted_avg = (grade_df.mul(pd.Series(grade_weights), axis=0).sum() / grade_df.loc["Total Appeared"]).round(2)
    weighted_avg.fillna(0, inplace=True)

    # Store "Average" as a separate DataFrame
    average_df = pd.DataFrame(weighted_avg).T  # Convert Series to DataFrame
    average_df.index = ["Average"]
    
    # Assign Grade from Average Score
    average_df.loc["Average"] = weighted_avg
    average_df.loc["Grade"] = weighted_avg.map(get_grade_from_average)

    # Merge "Average" and "Grade" rows separately 
    grade_df = pd.concat([grade_df, average_df])
    
    # Overall Average Calculation
    overall_avg = weighted_avg.mean().round(2)
    grade_df["Overall"] = ""
    grade_df.loc["Average", "Overall"] = overall_avg
    grade_df.loc["Grade", "Overall"] = get_grade_from_average(overall_avg)
    
    return grade_df

if __name__ == '__main__':
    file_path = r"Analyze Data\DDMCA_Sem7_2021.xlsx"
    df = pd.read_excel(file_path)
    grade_distribution_df = perform_analysis(df)
    print(grade_distribution_df)
