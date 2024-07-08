from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import os
import pandas as pd
from linkedin_api import Linkedin
from urllib.parse import urlparse
import json
import logging
from urllib.parse import unquote

app = Flask(__name__,template_folder='C:/Users/AdithiMadduluri/Desktop/contact scrapper')

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_profile_name(linkedin_url):
    parsed_url = urlparse(unquote(linkedin_url))
    path_segments = [
        segment for segment in parsed_url.path.split('/') if segment]
    company_index = path_segments.index('in') 
    if company_index < len(path_segments) - 1:
        return path_segments[company_index + 1] 
    if company_index < len(path_segments) - 1:
        return path_segments[company_index + 1]
    return None


def fetch_profile_details(api, profile_name, linkedin_url):
    try:
        profile_data = api.get_profile(profile_name)
        return json.loads(json.dumps(profile_data, indent=2))
    except Exception as e:
        logger.error(
            f"Error fetching details for {profile_name} from {linkedin_url}: {str(e)}")
        return None
def process_urls(api, linkedin_urls_df):
    all_profile_details = []

    for _, row in linkedin_urls_df.iterrows():
        linkedin_url = row['LinkedIn URL']
        profile_name = extract_profile_name(linkedin_url)
        if profile_name:
            try:
                prof_data = fetch_profile_details(api, profile_name, linkedin_url)                
                loc=""
                if 'geoLocationName' in prof_data.keys(): 
                    if prof_data['geoLocationName']:
                        loc = prof_data['geoLocationName'] 

                if prof_data:
                    extracted_data = {
                        "Linkedin URL": linkedin_url,
                        'FirstName': prof_data['firstName'],
                        'LastName': prof_data['lastName'],
                        'Main HeadLine': prof_data['headline'],
                        'Current Job': prof_data['experience'][0]['title'],
                        'Current company': prof_data['experience'][0]['companyName'],
                        'Country':prof_data['geoCountryName'],
                        'Location': loc
                    }
                    all_profile_details.append(extracted_data)
                else:
                    logger.warning(
                        f"No data fetched for {profile_name} from {linkedin_url}.")
                    extracted_data= {
                        "Linkedin URL": linkedin_url,
                        'FirstName': "",
                        'LastName': "",
                        'Main HeadLine': "",
                        'Current Job': "",
                        'Current company': "",
                        'Country': "",
                        'Location': ""
                    }
                    all_profile_details.append(extracted_data)
            except Exception as e:
                logger.warning(f"Error processing {linkedin_url}: {str(e)}")
                extracted_data = {
                    "Linkedin URL": linkedin_url,
                    'FirstName': "",
                    'LastName': "",
                    'Main HeadLine': "",
                    'Current Job': "",
                    'Current company': "",
                    'Country': "",
                    'Location': ""
                }
                all_profile_details.append(extracted_data)
                # If there's an error, add an empty row
        else:
            logger.warning(f"Invalid LinkedIn URL: {linkedin_url}")
            # If the URL is invalid, add an empty row
            extracted_data = {
                "Linkedin URL": linkedin_url,
                'FirstName': "",
                'LastName': "",
                'Main HeadLine': "",
                'Current Job': "",
                'Current company': "",
                'Country': "",
                'Location': ""
            }
            all_profile_details.append(extracted_data)

    return pd.DataFrame(all_profile_details)

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)

        file = request.files['file']

        if file.filename == '':
            return redirect(request.url)

        if file and allowed_file(file.filename):
            filename = os.path.join(app.config['UPLOAD_FOLDER'], 'urls.xlsx')
            file.save(filename)

            linkedin_email = request.form['email']
            linkedin_password = request.form['password']

            if not linkedin_password:
                return render_template('index.html', error='Password cannot be empty.')

            try:
                api = Linkedin(linkedin_email, linkedin_password)

                linkedin_urls_df = pd.read_excel(filename)
                processed_data = process_urls(api, linkedin_urls_df)

                output_filename = 'contacts_data_with_details.xlsx'
                output_filepath = os.path.join(
                    app.config['UPLOAD_FOLDER'], output_filename)

                processed_data.to_excel(output_filepath, index=False)

                return render_template('index.html', success=True, output_filename=output_filename)
            except Exception as e:
                error_message = f"Error: {str(e)}"

                logger.error(error_message)
                return render_template('index.html', error=error_message)

    return render_template('index.html', success=False)


if __name__ == '__main__':
    app.run(debug=True)
