import logging
import os
import time
import tkinter as tk
import warnings
from tkinter import filedialog, messagebox, scrolledtext

import numpy as np
import pandas as pd

from utils import *

warnings.filterwarnings("ignore")
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def processing(filepath=None, email=None):
    API17TRACK__KEY = os.environ.get("API17TRACK__KEY")
    GOOGLE_CREDS_JSON = "creds/google.json"

    logging.info("Welcome to our shopify export tracking dashboard!!")

    data = get_shopify_export()
    tracking_numbers = data["tracking_number"].tolist()
    tracking_numbers = list(set(tracking_numbers))
    print(f"Found {len(tracking_numbers)} unique tracking numbers.")

    df_country = pd.read_csv("data/country-codes.csv")
    country_mappinp = dict(zip(df_country["alpha-2"], df_country["country"]))
    del df_country

    print("Retrieving tracking data from 17track.net...")
    # Initialize counters
    registered_count = 0
    skipped_count = 0
    quota_exceeded_count = 0
    output = []
    for number in tracking_numbers[:400]:
        logging.debug(f"Retrieving data for {number}...")
        track = API17Track(API_KEY=API17TRACK__KEY)
        try:
            track.retrieve_package_data(number)
        except Exception as e:
            try:
                logging.debug(
                    f"{number} not registered. Registering & Retrieving now..."
                )
                track.register_package(number)
                registered_count += 1
            except API17TrackError as retry_e:
                logging.debug(
                    f"Retry failed for {number} with {retry_e}. Skipping to next."
                )
                if retry_e.code.code == track.QUOTA_LIMIT:
                    quota_exceeded_count += 1
                skipped_count += 1
                continue

        output.append(track.data)

    # Log the summary
    logging.info(f"Registration needed for {registered_count} package(s).")
    logging.info(f"Quota exceeded for {quota_exceeded_count} package(s).")
    logging.info(f"Skipped {skipped_count} package(s).")
    logging.info(f"Successfully retrieved data for {len(output)} package(s).")

    print("Processing data...")
    df = pd.DataFrame(output)
    df.shipping_country = df.shipping_country.apply(
        lambda x: get_country(country_mappinp, x)
    )
    df.recipient_country = df.recipient_country.apply(
        lambda x: get_country(country_mappinp, x)
    )
    df = df[
        [
            "tracking_number",
            "carrier_name",
            "shipping_country",
            "recipient_country",
            "latest_status",
            "days_after_order",
            "days_of_transit",
            "events.InTransit",
            "events.Delivered",
            "events.InfoReceived",
        ]
    ]
    df.rename(
        columns={
            "events.InTransit": "in_transit_at",
            "events.Delivered": "delivered_at",
            "events.InfoReceived": "info_received_at",
        },
        inplace=True,
    )
    data = pd.merge(data, df, how="left", on="tracking_number")

    # remove rows with latest_status null, or having Exception NotFound DeliveryFailure
    data = data[~data.latest_status.isnull()]
    data = data[
        ~data.latest_status.str.contains(
            "|".join(["DeliveryFailure", "NotFound", "Exception"])
        )
    ]

    data[
        [
            "processing_time",
            "shipping_time",
            "total_time",
            "order_created_at",
            "info_received_at",
            "in_transit_at",
            "delivered_at",
        ]
    ] = data.apply(get_shipping_metrics, axis=1, result_type="expand")

    print("Uploading data to Google Sheets...")
    IS_UPLOAD = True
    if IS_UPLOAD:
        spreadsheet_client = get_google_client_spreadsheet(GOOGLE_CREDS_JSON)
        drive_client = get_google_drive_client(GOOGLE_CREDS_JSON)
        spreadsheet_id = create_google_spreadsheet(
            spreadsheet_client,
            f"Shopify-Tracked-{email}-{time.time()}",
            sharedEmail=email,
            drive_client=drive_client,
        )
        upload_dataframe_to_google_sheet(
            data, "Sheet1", spreadsheet_client, spreadsheet_id
        )

        # print the google sheet url
        print(
            f"Upload Complete!\nSheet URL : https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid=0"
        )

    data.to_csv("data/output.csv", index=False)

    # print the dashboard url
    IS_DASHBOARD = True
    if IS_DASHBOARD:
        reportId = ""
        pageId_list = ["1lviD"]
        connector = "googleSheets"
        mode = "view"
        spreadsheetId = spreadsheet_id
        worksheetId = "0"
        alias = "ds0"

        dashboard_url = get_tracking_dashboard(
            reportId, pageId_list, connector, mode, spreadsheetId, worksheetId, alias
        )
        print(f"Dashboard URL : {dashboard_url}")
    return {
        "spreadsheet_url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid=0",
        "dashboard_url": dashboard_url,
    }


def main():
    root = tk.Tk()
    root.title("Shopify Tracking Dashboard")

    def open_file_dialog():
        file_path = filedialog.askopenfilename()
        if file_path:
            file_path_entry.delete(0, tk.END)
            file_path_entry.insert(0, file_path)

    def run_processing():
        file_path = file_path_entry.get()
        email = email_entry.get()
        if file_path:
            try:
                # Assuming process_file returns a dictionary with URLs
                urls = processing(file_path, email)
                urls_text_area.configure(state="normal")
                urls_text_area.delete(1.0, tk.END)

                # Format and insert the URLs with labels
                if "spreadsheet_url" in urls:
                    urls_text_area.insert(
                        tk.END, f"Google Sheet URL: {urls['spreadsheet_url']}\n\n"
                    )
                if "dashboard_url" in urls:
                    urls_text_area.insert(
                        tk.END, f"Looker Dashboard URL: {urls['dashboard_url']}\n"
                    )

                urls_text_area.configure(state="disabled")
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showwarning("Warning", "Please select a file first.")

    # File Path Entry
    file_path_label = tk.Label(root, text="Select your Shopify Template File:")
    file_path_label.pack()
    file_path_entry = tk.Entry(root, width=50)
    file_path_entry.pack()
    file_browse_button = tk.Button(root, text="Browse", command=open_file_dialog)
    file_browse_button.pack()
    email_label = tk.Label(root, text="Provide your email address:")
    email_label.pack()
    email_entry = tk.Entry(root, width=30)
    email_entry.pack()

    # Run Button
    run_button = tk.Button(
        root, text="Process Tracking Information", command=run_processing
    )
    run_button.pack()

    # URLs Text Area
    urls_text_area = scrolledtext.ScrolledText(root, state="disabled", height=10)
    urls_text_area.pack()

    # Set a function to run when closing the window
    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()
