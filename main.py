from decimal import DivisionUndefined
from numpy import nan
import numpy as np
import pandas as pd
import math
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from datetime import datetime

current_date = datetime.now()
formatted_date = current_date.strftime("%m-%d-%Y")
print(formatted_date)
print("\n")

df = pd.read_excel("report.xlsx", sheet_name="EMPLOYEE FOCUS", skiprows=5)


def get_data(name):
    get_name = df[df["Employee"] == name]
    employee = get_name.iloc[0, 1]

    # Check if the employee exists in the data
    if get_name.empty:
        print(f"Employee {name} not found.")
        return None

    # gp related variables
    gp_target = (get_name.iloc[0, 49])
    get_gp = math.floor(get_name.iloc[0, 50])  # Apply math.floor to the value
    get_trendingGP = math.floor(get_name.iloc[0, 51])

    get_pga_target = get_name.iloc[0, 19]
    get_PGA = get_name.iloc[0, 20]
    get_TrendingPGA = math.ceil(get_name.iloc[0, 21])

    # OGA Variables
    get_oga = int(get_name.iloc[0, 35])
    get_oga_target = int(get_name.iloc[0, 34])
    get_trendingOGA = int(math.ceil(get_name.iloc[0, 36]))

    get_protection = int(get_name.iloc[0, 40])
    get_ProtectionPercentage = math.floor(get_name.iloc[0, 43] * 100)
    trending_protection = round(get_name.iloc[0, 41], 2)

    # Accessory Variables
    Accesory_gp = get_name.iloc[0, 46]
    trending_accesory_gp = round(get_name.iloc[0, 47], 2)

    # Home Solutions Variables
    get_home_solutions = math.floor(get_name.iloc[0, 25])
    get_home_solutions_target = get_name.iloc[0, 23]
    get_trending_home_solutions = math.ceil(get_name.iloc[0, 26])

    # Misc Statistics
    total_boxes = get_name.iloc[0, 6]
    get_gp_per_invoice = round((get_name.iloc[0, 4]), 2)
    gp_per_box = round((get_name.iloc[0, 8]), 2)

    invoice_stats_dict = {
        "GP Per Invoice": get_gp_per_invoice,
        "GP per Box": gp_per_box,
        "Total Boxes Sold": total_boxes,
    }

    gp_dict = {
        "GP Target": gp_target,
        "Current GP": get_gp,
        "Trending GP": get_trendingGP,
    }
    pga_dict = {
        "PGA Target": get_pga_target,
        "PGA": get_PGA,
        "Trending PGA": get_TrendingPGA,
    }
    Oga_dict = {
        "OGA Target": get_oga_target,
        "OGA": get_oga,
        "Trending OGA": get_trendingOGA,     
    }

    home_dict = {
        "Home Solutions Target": get_home_solutions_target,
        "Home Solutions Sold": get_home_solutions,
        "Trending Home Solutions": get_trending_home_solutions,
    }

    vmp_dict = {
        "Protection": get_protection,
        "Trending Protection": trending_protection,
        "Protection Percentage": get_ProtectionPercentage
    }

    accesory_dict = {
        "Accesory GP": Accesory_gp,
        "Trending Accesory GP": trending_accesory_gp,
    }

    return {
        "INVOICE": invoice_stats_dict,
        "GP": gp_dict,
        "PGA": pga_dict,
        "OGA": Oga_dict,
        "HOME": home_dict,
        "VMP": vmp_dict,
        "ACC": accesory_dict,
    }


def createGraphs(employee, Category_data, Category_name, name):
    # Define categories with their corresponding data in the correct structure

    # Iterate through each category to generate graphs

    # Create x-axis labels for the bars (use keys of the category dictionary)
    x_axis = list(Category_data.keys())
    y_axis = list(Category_data.values())

    # Create a new figure for each category
    fig, ax = plt.subplots(figsize=(10, 8))

    # Plot the bar graph with the values
    ax.bar(x_axis, y_axis, color='red')

    # Add labels and title
    ax.set_xlabel(f'{Category_name} Statistics {formatted_date}')
    ax.set_ylabel(name)
    ax.set_title(f'{employee} - {Category_name} Statistics')

    # Rotate the x-axis labels to avoid overlap
    plt.xticks(rotation=45, ha="right")
    

    # Add the value on top of each bar
    for i in range(len(y_axis)):
        if Category_name not in ["Home", "Invoice", "PGA"]:
            ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'${x:,.2f}'))
            ax.text(i, y_axis[i] + 1, f'${y_axis[i]:,.2f}', ha='center', va='bottom')
        else:
            ax.text(i, y_axis[i] + 1, str(y_axis[i]), ha="center", va="bottom")

    # Display the graph
    plt.tight_layout()  # Adjust the layout to prevent clipping

    # Save the graph as a PNG file
    filename = f"{employee}_{Category_name}_{formatted_date}_graph.png"
    plt.savefig(filename)
    plt.show()
    
    

def DynamicGraphing(name, target_kpi):
    data = get_data(name)

    if target_kpi in data:
        stats = data[target_kpi]
        print(f"\n--- {target_kpi} ---")
        for metric, value in stats.items():
            print(f"{metric}: {value}")
        createGraphs(name, stats, target_kpi, target_kpi)
    else:
        print(f"KPI category '{target_kpi}' not found for {name}.")


full_name = input("Enter Employees *FULL NAME*: ").title()
kpi_input = input("Enter KPI Metric: ").upper()
if full_name is not None:
    if kpi_input is not None:
        DynamicGraphing(full_name, kpi_input)



    

    
    
    
            








