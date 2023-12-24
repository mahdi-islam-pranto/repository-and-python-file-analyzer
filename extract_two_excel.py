import requests
import os
from zipfile import ZipFile
import openpyxl
import re
from collections import Counter

def download_and_unzip_github_repository(repo_url, access_token):
    # Extract username and repository name from the GitHub URL
    _, _, _, username, repository = repo_url.rstrip('/').split('/')

    # Create a zip file name based on the repository name
    zip_file_name = f"{username}_{repository}_master.zip"

    # Construct the GitHub API URL to get the zipball of the repository
    api_url = f"https://api.github.com/repos/{username}/{repository}/zipball/master"

    # Headers for authentication with your personal access token
    headers = {
        'Authorization': f'token {access_token}'
    }

    # Send a GET request to the GitHub API to download the zipball
    response = requests.get(api_url, headers=headers)

    if response.status_code == 200:
        # Save the zipball to a local file
        with open(zip_file_name, 'wb') as zip_file:
            zip_file.write(response.content)

        print(f"Repository downloaded successfully as {zip_file_name}")

        # Unzip the downloaded file
        with ZipFile(zip_file_name, 'r') as zip_ref:
            zip_ref.extractall()

        print(f"Repository unzipped successfully.")

        # Analyze the unzipped repository, get commit information, and create Excel sheets
        analyze_and_create_excel(repo_url, access_token)
        analyze_python_files_and_create_excel()

    else:
        print(f"Failed to download repository. Status code: {response.status_code}")

def analyze_and_create_excel(repo_url, access_token):
    # Create a new Excel workbook and select the active sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Set header rows for repository information
    sheet.append(["Repository Information"])
    sheet.append(["Repo Name", "All Commitors", "Total Number of Commits"])

    print("\nAnalyzing the unzipped repository:")

    # Extract owner and repo names from the GitHub URL
    _, _, _, owner, repo = repo_url.rstrip('/').split('/')

    # GitHub API base URL
    api_url = 'https://api.github.com/'

    # Headers for authentication with your personal access token
    headers = {
        'Authorization': f'token {access_token}'
    }

    # Get the list of commits for the repository
    commits_url = f'{api_url}repos/{owner}/{repo}/commits'
    response = requests.get(commits_url, headers=headers)

    if response.status_code == 200:
        commits = response.json()
        repo_name = repo
        commitors = set()
        total_commits = len(commits)

        # Create a dictionary to track commit counts per developer
        commit_counts = {}

        for commit in commits:
            developer_name = commit['commit']['author']['name']
            developer_email = commit['commit']['author']['email']

            # Update the commit count for the developer
            if developer_name in commit_counts:
                commit_counts[developer_name] += 1
            else:
                commit_counts[developer_name] = 1

            commitors.add(developer_name)

        # Add repository information to the Excel sheet
        sheet.append([repo_name, ", ".join(commitors), total_commits])

        # Set header row for commit information
        sheet.append([])  # Empty row as a separator
        sheet.append(["Developers Information"])
        sheet.append(["Commiter's Name", "Committer's Email", "Number of Commits", "Commit Date and Time", "Commit Message"])

        for commit in commits:
            developer_name = commit['commit']['author']['name']
            developer_email = commit['commit']['author']['email']
            commit_date = commit['commit']['author']['date']
            commit_message = commit['commit']['message']

            # Add commit information to the Excel sheet
            sheet.append([developer_name, developer_email, commit_counts[developer_name], commit_date, commit_message])

        # Save the Excel workbook
        excel_file_name = "repository_analysis.xlsx"
        workbook.save(excel_file_name)

        print(f"\nExcel sheet created successfully: {excel_file_name}")

        # Print the number of commits per developer
        for developer, count in commit_counts.items():
            print(f"{developer} has {count} commits.")

    else:
        print(f"Error fetching commits: {response.status_code}")

def analyze_python_files_and_create_excel():
    # Additional code for Python file analysis
    print("\nAnalyzing Python files:")

    # Walk through the current directory and add Python file paths to the Excel sheet
    python_files = [os.path.join(root, file) for root, dirs, files in os.walk(".") for file in files if file.endswith(".py")]

    # Create a new Excel workbook for Python file analysis
    workbook_python = openpyxl.Workbook()
    sheet_python = workbook_python.active

    # Set header row for Python file analysis
    sheet_python.append(["Python File Analysis"])
    sheet_python.append(["File Name", "Comments", "Total Comments", "Class Names", "Method Names", "Comment Percentage", "Total Lines of Code", "ATFD"])

    for file_path in python_files:
        print(f"Analyzing Python file: {file_path}")
        file_name = os.path.basename(file_path)
        comments, comment_lines, methods, class_names, comment_percentage, total_lines = extract_comments_methods(file_path)

        # Get method names and class names
        method_names = [method[0] for method in methods]
        class_names_list = ', '.join(class_names).split(', ')

        # Count method occurrences
        method_counter = Counter(method_names)

        # Get method names used in multiple classes
        methods_used_in_multiple_classes = [method for method, count in method_counter.items() if count > 1]

        # Append data for the current file
        atfd = ', '.join([f"{method} ({method_counter[method]} times)" for method in methods_used_in_multiple_classes])
        sheet_python.append([file_name, '\n'.join(comments), comment_lines, ', '.join(class_names_list), ', '.join(method_names), comment_percentage, total_lines, atfd])

    # Save the Python file analysis Excel workbook
    excel_file_name_python = "python_file_analysis.xlsx"
    workbook_python.save(excel_file_name_python)

    print(f"\nExcel sheet for Python file analysis created successfully: {excel_file_name_python}")

# Additional code for extracting comments, methods, etc.
def extract_comments_methods(file_path):
    with open(file_path, 'r') as file:
        content = file.readlines()

    # Regular expressions to detect comments, methods, and class names
    comment_pattern = r'#.*'
    method_pattern = r'def\s+(\w+)\s*\((.*?)\):'
    class_pattern = r'class\s+(\w+)\s*:'

    # Find all comments in the file
    comments = []
    for line in content:
        if not line.strip():  # Skip empty lines
            continue
        match = re.match(comment_pattern, line.strip())
        if match:
            comments.append(match.group())

    # Find all methods and their parameters
    methods = re.findall(method_pattern, ''.join(content))

    # Find all class names
    class_names = re.findall(class_pattern, ''.join(content))

    # Calculate comment percentage and total lines (excluding empty lines)
    total_lines = len([line for line in content if line.strip()])
    comment_lines = len(comments)
    comment_percentage = (comment_lines / total_lines) * 100

    return comments, comment_lines, methods, class_names, comment_percentage, total_lines

if __name__ == "__main__":
    # Get GitHub repository URL and access token from the user
    github_url = input("Enter the GitHub repository URL: ")
    access_token = input("Enter your GitHub access token: ")

    # Download, unzip, and analyze the repository
    download_and_unzip_github_repository(github_url, access_token)
