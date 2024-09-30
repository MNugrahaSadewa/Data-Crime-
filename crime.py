import pandas as pd
import matplotlib.pyplot as plt

# Load the dataset from the Excel file
file_path = 'F:/Tel u/Data Sains/data Crime/crime dataset separate.xlsx'
df = pd.read_excel(file_path, sheet_name='crime dataset separate')

# Clean column names: remove spaces and special characters
df_cleaned = df.rename(columns=lambda x: x.strip().replace(' ', '_').replace('/', '_').replace('-', '_'))

# Drop columns that have too many missing values or seem irrelevant
columns_to_drop = ['Crm_Cd_3', 'Crm_Cd_4', 'Cross_Street', 'Mocodes', 'Weapon_Used_Cd', 'Weapon_Desc']
df_cleaned = df_cleaned.drop(columns=columns_to_drop)

# Check for missing values and drop rows if necessary
df_cleaned = df_cleaned.dropna(subset=['Crm_Cd_Desc', 'AREA_NAME'])

# Filter for the main relevant columns
relevant_columns = ['AREA_NAME', 'Vict_Age', 'Vict_Sex', 'Status_Desc', 'Crm_Cd_Desc', 'Premis_Desc', 'LOCATION']

# Create a filtered DataFrame with only relevant columns
df_filtered = df_cleaned[relevant_columns].dropna(how='all')  # Remove rows where all relevant columns are missing

# Function to visualize data based on user selection
def visualize_data(option):
    # Visualization options based on user input
    if option == 1:
        # Crime Distribution by Area Name
        area_distribution = df_cleaned['AREA_NAME'].value_counts().sort_index()
        plt.figure(figsize=(12, 8))  # Adjust figure size
        area_distribution.plot(kind='bar')
        plt.title('Crime Distribution by Area Name')
        plt.xlabel('Area Name')
        plt.ylabel('Count')
        plt.xticks(rotation=60, ha='right')  # Rotate x-axis labels for readability
        plt.tight_layout()
        plt.show()
    elif option == 2:
        # Victim Age Distribution (exclude age 0)
        victim_age_distribution = df_cleaned[df_cleaned['Vict_Age'] > 0]['Vict_Age'].value_counts().sort_index()
        plt.figure(figsize=(12, 8))  # Adjust figure size
        victim_age_distribution.plot(kind='bar')
        plt.title('Victim Age Distribution ')
        plt.xlabel('Age')
        plt.ylabel('Count')
        plt.xticks(rotation=60, ha='right')  # Rotate x-axis labels for readability
        plt.tight_layout()
        plt.show()
    elif option == 3:
        # Victim Sex Distribution
        victim_sex_distribution = df_cleaned['Vict_Sex'].value_counts()
        plt.figure(figsize=(8, 6))
        victim_sex_distribution.plot(kind='pie', autopct='%1.1f%%', startangle=140)
        plt.title('Victim Sex Distribution')
        plt.axis('equal')
        plt.tight_layout()
        plt.show()
    elif option == 4:
         # Crime Status Distribution (Bar Chart)
        crime_status_distribution = df_cleaned['Status_Desc'].value_counts()
        plt.figure(figsize=(12, 8))  # Adjust figure size
        crime_status_distribution.plot(kind='bar', color='skyblue')
        plt.title('Crime Status Distribution')
        plt.xlabel('Crime Status')
        plt.ylabel('Count')
        plt.xticks(rotation=45, ha='right')  # Rotate x-axis labels for readability
        plt.tight_layout()
        plt.show()
    elif option == 5:
        # Top 5 Crime Description (Crm_Cd_Desc) Distribution
        crm_cd_desc_distribution = df_cleaned['Crm_Cd_Desc'].value_counts().head(5).sort_index()  # Top 5
        plt.figure(figsize=(12, 8))  # Adjust figure size
        crm_cd_desc_distribution.plot(kind='bar')
        plt.title('Top 5 Crime Description Distribution ')
        plt.xlabel('Crime Description')
        plt.ylabel('Count')
        plt.xticks(rotation=60, ha='right')  # Rotate x-axis labels for readability
        plt.tight_layout()
        plt.show()
    elif option == 6:
        # Top 5 Premises Description (Premis_Desc) Distribution
        premis_desc_distribution = df_cleaned['Premis_Desc'].value_counts().head(5).sort_index()  # Top 5
        plt.figure(figsize=(14, 8))  # Increase figure size for better label readability
        premis_desc_distribution.plot(kind='bar')
        plt.title('Top 5 Premises Description Distribution ')
        plt.xlabel('Premises Description')
        plt.ylabel('Count')
        plt.xticks(rotation=60, ha='right')  # Rotate x-axis labels for better readability
        plt.tight_layout()  # Ensure layout fits labels properly
        plt.show()
    elif option == 7:
        # Location Distribution - Check for missing or incomplete names
        # Display top 10 locations
        location_distribution = df_cleaned['LOCATION'].dropna().value_counts().head(10)
        
        # Visualize top 10 locations with adjustments to display full names
        plt.figure(figsize=(12, 8))  # Adjust figure size
        location_distribution.plot(kind='bar')
        plt.title('Top 10 Locations with Most Crimes')
        plt.xlabel('Location')
        plt.ylabel('Count')
        plt.xticks(rotation=60, ha='right')  # Rotate labels for better readability
        plt.tight_layout()  # Adjust layout to avoid label cut-off
        plt.show()
    else:
        print("Invalid option selected.")

# Function to handle user input for visualization option
def get_visualization_option():
    while True:
        try:
            print("\nChoose a visualization option:")
            print("1. Crime Distribution by Area Name")
            print("2. Victim Age Distribution")
            print("3. Victim Sex Distribution")
            print("4. Crime Status Distribution")
            print("5. Top 5 Crime Description Distribution")
            print("6. Top 5 Premises Description Distribution")
            print("7. Top 10 Locations with Most Crimes")
            
            # Get the user's choice for visualization
            option = int(input("Enter the number of the option you want to visualize (1-7): "))
            
            # Validate the option
            if option in [1, 2, 3, 4, 5, 6, 7]:
                return option
            else:
                print("Invalid option. Please choose a number between 1 and 7.")
        except ValueError:
            print("Invalid input. Please enter a valid number (1-7).")

# Function to ask if the user wants to visualize again
def ask_to_continue():
    while True:
        cont = input("Do you want to visualize another option? (yes/no): ").lower()
        if cont in ['yes', 'no']:
            return cont == 'yes'
        else:
            print("Invalid input. Please enter 'yes' or 'no'.")

# Main flow of the program
continue_visualization = True

while continue_visualization:
    visualization_option = get_visualization_option()  # Get the user's visualization choice
    visualize_data(visualization_option)  # Call the function to visualize data
    
    # Save the filtered data to an Excel file after any visualization option
    output_file_path = 'F:/Tel u/Data Sains/data Crime/crime_analysis_output_filtered.xlsx'
    with pd.ExcelWriter(output_file_path) as writer:
        df_filtered.to_excel(writer, sheet_name='Filtered_Crime_Data', index=False)
    
    print(f"Filtered data saved to {output_file_path}")
    
    continue_visualization = ask_to_continue()  # Ask if the user wants to visualize again

