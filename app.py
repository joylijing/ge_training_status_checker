import os
from flask import Flask, request, render_template, send_file
import pandas as pd
from fuzzywuzzy import fuzz
from io import BytesIO

app = Flask(__name__)

def fuzzy_name_match(name1, name2):
    """Perform fuzzy name matching using FuzzyWuzzy."""
    return fuzz.token_sort_ratio(name1, name2) >= 80  # Adjust threshold as needed

def check_guest_editors(new_ge_df, tracker_df, archive_df):
    results = []
    for _, new_ge_row in new_ge_df.iterrows():
        name = new_ge_row['Guest Editor Name']
        email = new_ge_row['Email Address']
        found_in_tracker = False
        found_in_archive = False

        # Round 1: Exact email match
        if email in tracker_df['Email Address'].values:
            found_in_tracker = True
        if email in archive_df['Email Address'].values:
            found_in_archive = True

        # Round 2: Fuzzy name match if email not found
        if not found_in_tracker and not found_in_archive:
            for _, tracker_row in tracker_df.iterrows():
                if fuzzy_name_match(name, tracker_row['Guest Editor Name']):
                    found_in_tracker = True
                    break
            for _, archive_row in archive_df.iterrows():
                if fuzzy_name_match(name, archive_row['Guest Editor Name']):
                    found_in_archive = True
                    break

        # Determine Training Status and Included In
        training_status = "Yes" if found_in_tracker or found_in_archive else "No"
        included_in = []
        if found_in_tracker:
            included_in.append("Tracker")
        if found_in_archive:
            included_in.append("Archive")
        included_in = ", ".join(included_in) if included_in else ""

        # Record results
        results.append({
            'Guest Editor Name': name,
            'Email Address': email,
            'Training Status': training_status,
            'Included In': included_in
        })

    return results

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    # Get uploaded files
    tracker_file = request.files['tracker']
    archive_file = request.files['archive']
    new_ge_file = request.files['new_ge']

    # Load data into DataFrames using openpyxl as the engine
    tracker_df = pd.read_excel(tracker_file, engine='openpyxl')
    archive_df = pd.read_excel(archive_file, engine='openpyxl')
    new_ge_df = pd.read_excel(new_ge_file, engine='openpyxl')

    # Perform checks
    results = check_guest_editors(new_ge_df, tracker_df, archive_df)

    # Convert results to DataFrame
    results_df = pd.DataFrame(results)

    # Save the updated DataFrame to an Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        results_df.to_excel(writer, index=False)
    output.seek(0)

    # Return the Excel file as a downloadable response
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Updated_New_GE.xlsx'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 10000)))