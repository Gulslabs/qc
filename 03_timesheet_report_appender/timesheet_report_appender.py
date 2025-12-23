import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import sys

def parse_time(time_str):
    """Convert time string (HH:MM) to datetime object"""
    try:
        return datetime.strptime(str(time_str).strip(), "%H:%M")
    except:
        return None

def time_to_decimal(dt):
    """Convert datetime to decimal (Excel format)"""
    if pd.isna(dt) or dt is None:
        return None
    return (dt.hour * 3600 + dt.minute * 60 + dt.second) / 86400

def decimal_to_time_str(decimal):
    """Convert decimal back to time string"""
    if pd.isna(decimal):
        return ""
    total_seconds = decimal * 86400
    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    return f"{hours:02d}:{minutes:02d}"

def create_time_bins(min_time, max_time, interval_minutes):
    """Create time bins from min to max with given interval"""
    bins_start = []
    bins_end = []
    
    current = min_time
    cutoff_time = datetime.strptime("23:00", "%H:%M")
    
    while current <= max_time:
        next_time = current + timedelta(minutes=interval_minutes)
        bins_start.append(current)
        bins_end.append(next_time)
        current = next_time
        
        # Stop at first occurrence after 23:00
        if next_time > cutoff_time:
            break
    
    return bins_start, bins_end

def calculate_frequency(times, bin_ends):
    """Calculate frequency distribution like Excel FREQUENCY function"""
    counts = []
    time_decimals = [time_to_decimal(t) for t in times if t is not None]
    bin_end_decimals = [time_to_decimal(b) for b in bin_ends]
    
    for i, bin_end in enumerate(bin_end_decimals):
        if i == 0:
            # First bin: count values <= bin_end
            count = sum(1 for t in time_decimals if t <= bin_end)
        else:
            # Other bins: count values > previous bin_end and <= current bin_end
            count = sum(1 for t in time_decimals if bin_end_decimals[i-1] < t <= bin_end)
        counts.append(count)
    
    return counts

def process_excel_file(file_path, interval_minutes=5):
    """
    Process Excel file and add time bin analysis columns
    
    Parameters:
    - file_path: Path to the Excel file
    - interval_minutes: Time interval for bins (default: 5 minutes)
    """
    
    # Read the Excel file
    # Row 1-2: Metadata, Row 3: Headers, Row 4+: Data
    print(f"Reading Excel file: {file_path}")
    df = pd.read_excel(file_path, header=2)  # Row 3 (index 2) is the header
    
    # Find the last row with data
    last_row = df.dropna(subset=['Name']).index[-1] + 1  # +1 for 0-indexing
    print(f"Data found up to row: {last_row + 1}")  # +1 for Excel row number
    
    # Parse Join Time and Leave Time
    print("Parsing time columns...")
    join_times = [parse_time(t) for t in df['Join Time'].dropna()]
    leave_times = [parse_time(t) for t in df['Leave Time'].dropna()]
    
    # Filter out None values
    join_times = [t for t in join_times if t is not None]
    leave_times = [t for t in leave_times if t is not None]
    
    if not join_times or not leave_times:
        print("Error: No valid time data found!")
        return
    
    # Find min and max times
    min_join = min(join_times)
    max_join = max(join_times)
    min_leave = min(leave_times)
    max_leave = max(leave_times)
    
    print(f"Join time range: {min_join.strftime('%H:%M')} to {max_join.strftime('%H:%M')}")
    print(f"Leave time range: {min_leave.strftime('%H:%M')} to {max_leave.strftime('%H:%M')}")
    
    # Create time bins for Join Time
    print(f"\nCreating {interval_minutes}-minute bins for Join Time...")
    join_bins_start, join_bins_end = create_time_bins(min_join, max_join, interval_minutes)
    
    # Create time bins for Leave Time
    print(f"Creating {interval_minutes}-minute bins for Leave Time...")
    leave_bins_start, leave_bins_end = create_time_bins(min_leave, max_leave, interval_minutes)
    
    # Calculate frequencies
    print("Calculating student counts...")
    join_counts = calculate_frequency(join_times, join_bins_end)
    leave_counts = calculate_frequency(leave_times, leave_bins_end)
    
    # Create new columns in DataFrame
    max_bins = max(len(join_bins_start), len(leave_bins_start))
    
    # Initialize columns with empty values
    df['Bin_Start_Join_Time'] = ''
    df['Bin_End_Join_Time'] = ''
    df['Join_Student_Count'] = ''
    df['Bin_Start_Leave_Time'] = ''
    df['Bin_End_Leave_Time'] = ''
    df['Leave_Student_Count'] = ''
    
    # Fill in the bin data starting from row 0 in DataFrame (which is row 4 in Excel)
    # Since we read with header=2, the DataFrame starts fresh from 0
    start_row = 0  # First data row in DataFrame (Row 4 in Excel)
    
    for i in range(len(join_bins_start)):
        if start_row + i < len(df):
            df.at[start_row + i, 'Bin_Start_Join_Time'] = join_bins_start[i].strftime('%H:%M')
            df.at[start_row + i, 'Bin_End_Join_Time'] = join_bins_end[i].strftime('%H:%M')
            df.at[start_row + i, 'Join_Student_Count'] = join_counts[i]
    
    for i in range(len(leave_bins_start)):
        if start_row + i < len(df):
            df.at[start_row + i, 'Bin_Start_Leave_Time'] = leave_bins_start[i].strftime('%H:%M')
            df.at[start_row + i, 'Bin_End_Leave_Time'] = leave_bins_end[i].strftime('%H:%M')
            df.at[start_row + i, 'Leave_Student_Count'] = leave_counts[i]
    
    # Save to new Excel file
    output_file = file_path.replace('.xlsx', '_with_report.xlsx')
    print(f"\nSaving results to: {output_file}")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Time_Sheet_Curated', index=False, header=True, startrow=2)
    
    # Print summary
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print(f"\nJoin Time Bins (Total: {len(join_bins_start)}):")
    print(f"{'Time Range':<20} {'Student Count':<15}")
    print("-" * 35)
    for i in range(len(join_bins_start)):
        time_range = f"{join_bins_start[i].strftime('%H:%M')}-{join_bins_end[i].strftime('%H:%M')}"
        print(f"{time_range:<20} {join_counts[i]:<15}")
    
    print(f"\nLeave Time Bins (Total: {len(leave_bins_start)}):")
    print(f"{'Time Range':<20} {'Student Count':<15}")
    print("-" * 35)
    for i in range(len(leave_bins_start)):
        time_range = f"{leave_bins_start[i].strftime('%H:%M')}-{leave_bins_end[i].strftime('%H:%M')}"
        print(f"{time_range:<20} {leave_counts[i]:<15}")
    
    print(f"\n✓ Processing complete! Output saved to: {output_file}")
    return df

# Main execution
if __name__ == "__main__":
    # Example usage
    if len(sys.argv) < 2:
        print("Usage: python script.py <excel_file_path> [interval_minutes]")
        print("Example: python script.py attendance.xlsx 5")
        print("\nOr modify the script and run directly:")
        print("  file_path = 'your_file.xlsx'")
        print("  interval_minutes = 5")
        sys.exit(1)
    
    file_path = sys.argv[1]
    interval_minutes = int(sys.argv[2]) if len(sys.argv) > 2 else 5    
    process_excel_file(file_path, interval_minutes)