# Result Analysis Tool

This is a Flask-based web application designed for result analysis. It provides functionalities to register users, perform result analysis, and generate reports. The application uses SQLAlchemy for database management, Flask-WTF for form handling, and Flask-Login for user authentication.

## Prerequisites

Before running the application, make sure you have the required Python libraries installed. You can install them using the following command:

```bash
pip install requirement.txt
```

## Getting Started

1. Clone the repository to your local machine:

```bash
git clone https://github.com/satishkumarmurugan/Student-Result-Generators-and-Analyzers.git 
```

2. Run the application:

```bash
python app.py
```

## Usage

### User Registration

- Access the application at `http://localhost:5000/` in your web browser.
- Click on the "Register" link in the navigation bar.
- Fill out the registration form with your name, email, and password.
- Click on the "Register" button to create an account.

### User Login

- Click on the "Login" link in the navigation bar.
- Enter your registered email and password.
- Click on the "Login" button.

### Result Analysis

- Navigate to the "Result Analysis" section.
- Upload an Excel file containing student result data.
- Perform various result analysis operations, such as counting failures, calculating pass percentages, etc.
- View the result analysis report.

### Comparison Tool

- Navigate to the "Comparison Tool" section.
- Upload Excel files for comparison.
- The tool will merge the data and extract relevant columns for further analysis.
- View the formatted output in a new Excel file.

### Topper Analysis

- Navigate to the "Topper Analysis" section.
- Upload an Excel file containing student result data.
- The application will identify the top scorers and generate an output Excel file.

### Logout

- Click on the "Logout" link in the navigation bar to log out of your account.

## Important Notes

- Make sure to set a secure secret key (`app.secret_key`) and database URI (`app.config['SQLALCHEMY_DATABASE_URI']`) before deploying the application in a production environment.
- This readme assumes basic knowledge of Flask and web application deployment.

Feel free to customize and extend the application based on your specific requirements!
