import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import platform
import tkinter as tk
from tkinter import filedialog

# Define constants and location coordinates
latitude = 7.0707
longitude = 125.6113
elevation = 7  # meters
solar_constant = 1361  # Updated solar constant in W/m²

def browse_save_file():
    """Open file dialog for selecting save location and filename"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.asksaveasfilename(
        title="Save Processed File As",
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    return file_path

def compute_day_of_year(date_time):
    return date_time.timetuple().tm_yday

def compute_hour_of_day(date_time):
    """
    Compute hour of day based on End Period.
    For 0:00, returns 24. For all other times, returns hour + minute/60
    """
    if date_time.hour == 0 and date_time.minute == 0:
        return 24
    return date_time.hour + date_time.minute / 60

def compute_declination(day_of_year):
    return np.radians(23.45 * np.sin(np.radians(360 / 365 * (day_of_year + 284))))

def compute_solar_zenith_angle(day_of_year, hour_of_day):
    declination = compute_declination(day_of_year)
    hour_angle = np.radians(15 * (hour_of_day - 12))  # Hour angle in radians
    latitude_rad = np.radians(latitude)

    # Compute solar elevation angle
    solar_elevation_angle = np.arcsin(
        np.sin(latitude_rad) * np.sin(declination) +
        np.cos(latitude_rad) * np.cos(declination) * np.cos(hour_angle)
    )

    # Compute solar zenith angle
    solar_zenith_angle = np.pi / 2 - solar_elevation_angle

    return np.degrees(solar_zenith_angle)  # Solar zenith angle (°)

def compute_ghi_lags(data):
    """Compute GHI_lag (t-1)."""
    data['GHI_lag (t-1)'] = data['GHI - W/m^2'].shift(1)
    return data

def add_day_period_columns(data):
    """Add Daytime column based on Hour of Day."""
    data['Daytime'] = np.where((data['Hour of Day'] >= 6) & (data['Hour of Day'] <= 18), 1, 0)
    return data

def compute_month_of_year(date_time):
    """
    Compute month of year from timestamp with quarter-month precision.
    Returns: month as float (e.g., Jan = 1.0, 1.25, 1.5, 1.75; Feb = 2.0, 2.25, 2.5, 2.75)
    """
    month = date_time.month
    day = date_time.day
    
    # Determine which quarter of the month
    if day <= 7:
        quarter = 0.0
    elif day <= 14:
        quarter = 0.25
    elif day <= 21:
        quarter = 0.5
    else:
        quarter = 0.75
        
    return month + quarter

def determine_season(month):
    """
    Determine the season based on the month number.
    Args:
        month (float): Month number with quarter precision (e.g., 1.0, 1.25, etc.)
    Returns:
        int: Season category (1: Cool Dry, 2: Hot Dry, 3: Rainy)
    """
    base_month = int(month)  # Get the base month number without the quarter
    
    if base_month in [12, 1, 2]:
        return 1  # Cool Dry
    elif base_month in [3, 4, 5]:
        return 2  # Hot Dry
    else:  # months 6-11
        return 3  # Rainy

def process_weather_data():
    """
    Process weather data CSV file to calculate hourly averages and add derived features.
    Automatically finds dataset.csv in the current directory.
    """
    # Define input file - look for dataset.csv in current directory
    input_file = "dataset.csv"
    
    if not os.path.exists(input_file):
        print(f"'{input_file}' not found in current directory.")
        return
    
    print(f"Processing {input_file}...")
    
    # Read the CSV file with utf-8 encoding
    try:
        df = pd.read_csv(input_file, encoding='utf-8')
    except UnicodeDecodeError:
        # If UTF-8 fails, try with 'latin-1' or 'cp1252' encoding
        df = pd.read_csv(input_file, encoding='latin-1')
    
    print(f"Columns found in dataset: {df.columns.tolist()}")
    
    # Check if we already have the processed format (this might be a re-run on processed data)
    if 'Date' in df.columns and 'Start Period' in df.columns and 'End Period' in df.columns:
        print("Dataset appears to be already processed. Using existing format.")
        result_df = df
    else:
        # Look for a date/time column - try common variations
        date_time_cols = [col for col in df.columns if any(term in col.lower() for term in ['date', 'time'])]
        
        if not date_time_cols:
            print("Error: No date or time column found in the dataset")
            print("Available columns:", df.columns.tolist())
            return None
        
        # Use the first date/time column found
        date_time_col = date_time_cols[0]
        print(f"Using '{date_time_col}' as the date & time column")
        
        # Fix column names with degree symbol
        fixed_columns = {}
        for col in df.columns:
            if 'Temp - ' in col and 'Â°C' in col:
                fixed_columns[col] = 'Temp - °C'
            elif 'Dew Point - ' in col and 'Â°C' in col:
                fixed_columns[col] = 'Dew Point - °C'
            elif 'Wet Bulb - ' in col and 'Â°C' in col:
                fixed_columns[col] = 'Wet Bulb - °C'
        
        # Apply the column name fixes
        df = df.rename(columns=fixed_columns)
        
        # Rename Solar Rad column to GHI
        solar_col = [col for col in df.columns if 'Solar Rad' in col]
        if solar_col:
            df = df.rename(columns={solar_col[0]: 'GHI - W/m^2'})
        
        try:
            # Try to parse the date column with flexible format detection
            df['datetime'] = pd.to_datetime(df[date_time_col])
        except:
            print(f"Warning: Could not parse '{date_time_col}' as datetime. Using row index instead.")
            # Create a datetime column starting from today if parsing fails
            start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            df['datetime'] = [start_date + timedelta(minutes=i*10) for i in range(len(df))]
        
        # Create hourly groups
        df['hour_group'] = df['datetime'].dt.floor('H')
        
        # Identify columns to average (all except date/time columns)
        # FIXED: Exclude string columns and the datetime columns we created
        numeric_columns = []
        for col in df.columns:
            if col not in ['datetime', 'hour_group', date_time_col]:
                # Check if column contains numeric data
                try:
                    # Try to convert first value to float to test if numeric
                    float(df[col].iloc[0])
                    numeric_columns.append(col)
                except (ValueError, TypeError):
                    print(f"Skipping non-numeric column: {col}")
        
        print(f"Numeric columns that will be averaged: {numeric_columns}")
        
        # Calculate hourly averages
        hourly_data = []
        
        for hour, group in df.groupby('hour_group'):
            # Format date as "7-Mar-24" with Windows compatibility
            if platform.system() == 'Windows':
                # On Windows, remove leading zero manually
                day = hour.day
                month = hour.strftime("%b")
                year = hour.strftime("%y")
                date_str = f"{day}-{month}-{year}"
            else:
                # On Unix/Linux/Mac
                date_str = hour.strftime("%-d-%b-%y")
            
            # Format start period as "0:00:00"
            start_period = hour.strftime("%H:%M:%S")
            
            # Calculate end period (1 hour later)
            end_hour = hour + timedelta(hours=1)
            end_period = end_hour.strftime("%H:%M:%S")
            
            row = {
                'Date': date_str,
                'Start Period': start_period,
                'End Period': end_period
            }
            
            # Calculate average for each measurement column
            for column in numeric_columns:
                try:
                    if 'Wind Run' in column:
                        # For Wind Run, calculate the sum for the hour
                        row[column] = group[column].sum()
                    else:
                        # For all other measurements, calculate the average
                        row[column] = group[column].mean()
                except Exception as e:
                    print(f"Error processing column {column}: {e}")
                    row[column] = None  # Use None for failed calculations
                
            hourly_data.append(row)
        
        # Create result dataframe
        result_df = pd.DataFrame(hourly_data)
        
        # Extract all column names for reordering
        all_columns = result_df.columns.tolist()
        
        # Find UV Index and GHI columns
        uv_col = [col for col in all_columns if 'UV Index' in col]
        ghi_col = [col for col in all_columns if 'GHI' in col and 'lag' not in col.lower()]
        
        # Start with the date and time columns
        new_column_order = ['Date', 'Start Period', 'End Period']
        
        # Remove UV Index and GHI columns from the list as we'll add them at the end
        if uv_col:
            for col in uv_col:
                all_columns.remove(col) if col in all_columns else None
        
        if ghi_col:
            for col in ghi_col:
                all_columns.remove(col) if col in all_columns else None
        
        # Add all other columns (excluding those we've already handled)
        remaining_columns = [col for col in all_columns if col not in ['Date', 'Start Period', 'End Period']]
        new_column_order.extend(remaining_columns)
        
        # Add UV Index as second-to-last column
        if uv_col:
            new_column_order.extend(uv_col)
        
        # Add GHI columns as the very last
        if ghi_col:
            new_column_order.extend(ghi_col)
        
        # Reorder the DataFrame columns
        result_df = result_df[new_column_order]
        
        # Add additional parameters
        try:
            # Convert Date column to datetime (using the format we created: D-MMM-YY)
            result_df['Date_dt'] = pd.to_datetime(result_df['Date'], format='%d-%b-%y')
            
            # Compute date-based parameters using Date_dt column
            result_df['Day of Year'] = result_df['Date_dt'].apply(compute_day_of_year)
            result_df['Month of Year'] = result_df['Date_dt'].apply(compute_month_of_year)
            
            # Compute Hour of Day using End Period column
            result_df['Hour of Day'] = pd.to_datetime(result_df['End Period'], format='%H:%M:%S').dt.hour
            # Handle midnight (00:00) case
            result_df.loc[result_df['Hour of Day'] == 0, 'Hour of Day'] = 24
            
            # Compute solar zenith angle
            result_df['Solar Zenith Angle'] = result_df.apply(
                lambda row: compute_solar_zenith_angle(row['Day of Year'], row['Hour of Day']),
                axis=1
            )
            
            # Add Season column
            result_df['Season'] = result_df['Month of Year'].apply(determine_season)
            
            # Add GHI lags and day period columns
            result_df = compute_ghi_lags(result_df)
            result_df = add_day_period_columns(result_df)
            
            # Drop only the temporary datetime column
            result_df = result_df.drop('Date_dt', axis=1, errors='ignore')
            
            # Also drop Season if needed
            if 'Season' in result_df.columns:
                result_df = result_df.drop('Season', axis=1)
            
        except Exception as e:
            print(f"Error adding additional parameters: {str(e)}")
        
        # Save to CSV with UTF-8 encoding, overwriting the original dataset.csv
        result_df.to_csv('dataset.csv', index=False, encoding='utf-8')
        
        print(f"Processed data saved to: dataset.csv")
        
        # Display summary
        print(f"Processed {len(df)} measurements into {len(result_df)} hourly records")
        return result_df

if __name__ == "__main__":
    # Process the dataset
    hourly_df = process_weather_data()
    
    # Display the first few rows of results
    if hourly_df is not None:
        print("\nSample of processed data:")
        print(hourly_df.head())
