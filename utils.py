import pandas as pd
import numpy as np
import os
import re
import gspread
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from requests import get, post
from types import SimpleNamespace
import tkinter as tk
from tkinter import filedialog

"""
    Notes:
    Since we are using a service account, the file created by it will be only accessible to that account.
    So, we create the sheets using Drive API to add permission to the sheet, then use sheets API to edit the content.

"""

# ---------- 17Track API Helpers -----------
# 17Track API Docs: https://api.17track.net/en/doc?anchor=track-v2-2


class API17Track:
    API_BASE: str = "https://api.17track.net/track/v2.2/"
    TRACKING_REGISTERED = -18019901
    TRACKING_NEED_REGISTER = -18019902
    QUOTA_LIMIT = -18019908

    def __init__(self, API_KEY: str) -> None:
        self.data = {}
        self.api_key = API_KEY

    def _build_request(self, endpoint: str, payload: list):
        response = (
            post(
                url=f"{self.API_BASE + endpoint}",
                json=payload,
                headers={"content-type": "application/json", "17token": self.api_key},
            )
            .json(object_hook=lambda d: SimpleNamespace(**d))
            .data
        )  # Convert to Object and return the 'data' key

        if response.accepted:
            return response.accepted[0]

        match response.rejected[0].error.code:
            case self.TRACKING_REGISTERED | self.TRACKING_NEED_REGISTER:
                raise Exception(response.rejected[0].error.message)  # Default Throw
            case self.QUOTA_LIMIT | _:
                raise API17TrackError(
                    response.rejected[0].error, response.rejected[0].error.message
                )

    def register_package(self, tracking_number: str):
        """Register a Tracking on 17Track"""

        creation = self._build_request(
            "register",
            payload=[
                {
                    "number": tracking_number,
                }
            ],
        )

        return self.retrieve_package_data(
            tracking_number
        )  # Update the Order Data (Status, etc)

    def retrieve_package_data(self, tracking_number: str) -> dict:
        """Retrieve data from 17Track Tracking"""

        try:
            obj_17track = self._build_request(
                "gettrackinfo", payload=[{"number": tracking_number}]
            )
        except Exception as e:
            raise Exception(e)

        return self._build_order_data(obj_17track)

    def _build_order_data(self, obj_17track) -> dict:
        """Build our Object(Order Document) from 17Track Data."""

        self.data: dict = {
            "tracking_number": obj_17track.number,
            "carrier_name": obj_17track.track_info.tracking.providers[0].provider.name,
            "shipping_country": obj_17track.track_info.shipping_info.shipper_address.country,
            "recipient_country": obj_17track.track_info.shipping_info.recipient_address.country,
            "latest_status": obj_17track.track_info.latest_status.status,
            "days_after_order": obj_17track.track_info.time_metrics.days_after_order,
            "days_of_transit": obj_17track.track_info.time_metrics.days_of_transit,
        }

        if (hasattr(obj_17track.track_info.tracking.providers[0], "events")) and (
            obj_17track.track_info.tracking.providers[0].events is not None
        ):
            for event in obj_17track.track_info.tracking.providers[0].events:
                if hasattr(event, "sub_status") and hasattr(event, "time_raw"):
                    if "_" not in event.sub_status:
                        self.data[f"events.{event.sub_status}"] = event.time_raw.date
                    else:
                        sub_status = event.sub_status.split("_")[0]
                        self.data[f"events.{sub_status}"] = event.time_raw.date

        return self.data


class API17TrackError(Exception):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


# ---------- Google Sheet & Drive helpers -----------


def get_google_drive_client(creds_path):
    """get google drive client

    Args:
        creds_path (str): service account credentials file path

    Returns:
        _type_: client
    """

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    credentials = Credentials.from_service_account_file(creds_path, scopes=scopes)
    client = build("drive", "v3", credentials=credentials)

    return client


def get_google_client_spreadsheet(creds_path):
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    credentials = Credentials.from_service_account_file(creds_path, scopes=scopes)
    client = build("sheets", "v4", credentials=credentials)

    return client


def create_google_spreadsheet(
    spreadsheet_client,
    spreadsheet_name,
    folder_id="1HpDBpJ8W4f7EvLz3tGi3CWnjB0SSKgwD",
    drive_client=None,
):
    """Create a spreadsheet at the designated drive folder

    Args:
        spreadsheet_client (object): sheet client
        spreadsheet_name (str): name of the spreadsheet
        folder_id (str, optional): drive folder id. Defaults to '1HpDBpJ8W4f7EvLz3tGi3CWnjB0SSKgwD'.
        drive_client (object, optional): drive client. Defaults to None.

    Returns:
        _type_: _description_
    """
    spreadsheet = {
        "properties": {"title": spreadsheet_name},
    }
    response = spreadsheet_client.spreadsheets().create(body=spreadsheet).execute()

    spreadsheet_id = response["spreadsheetId"]
    drive_client.files().update(
        fileId=spreadsheet_id, addParents=folder_id, removeParents="root"
    ).execute()

    return spreadsheet_id


def upload_dataframe_to_google_sheet(
    dataframe,
    sheet_name,
    client,
    spreadsheet_id,
):
    """upload the dataframe to the google sheet

    Args:
        dataframe (pd.DataFrame): dataframe to be uploaded
        sheet_name (str): name of the worksheet. Note: this is not the name of the spreadsheet
        client (object): google sheet client
        spreadsheet_id (str): spreadsheet id
    """
    if not dataframe.empty:
        # We need to replace NaNs and Infs with None and "Infinity" because gsheets doesn't support them.
        df = dataframe.replace([np.inf, -np.inf], "Infinity")
        values = [df.columns.tolist()] + df.to_numpy(na_value=None).tolist()
        body = {"values": values}
        request_range = f"{sheet_name}!A1:{chr(65 + df.shape[1] - 1)}{df.shape[0] + 1}"
        request = (
            client.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                valueInputOption="RAW",
                range=request_range,
                body=body,
            )
        )
        request.execute()


# ---------- LookerStudio Dashboard Helpers -----------


def get_tracking_dashboard(
    reportId,
    pageId_list,
    connector,
    mode,
    spreadsheet_id,
    worksheet_id,
    alias,
    pagination=False,
):
    """Create a new dashboard given the template lookerstudio report and data source.

    Args:
        reportId (str): lookerstudio report id for the template dashboard
        pageId_list (str): lookerstudio report page is lists
        connector (str): data source connector
        mode (str): view/edit mode
        spreadsheet_id (str): data source spreadsheet id - google sheet
        worksheet_id (str): data source worksheet id - google sheet
        alias (str): data source alias
        pagination (bool, optional): Whether report has multiple pages with different data sources. Defaults to False.
    """

    if pagination:
        for pageId in pageId_list:
            # currently out of scope
            url = (
                "https://lookerstudio.google.com/reporting/create?"
                "c.reportId={reportId}&"
                "c.pageId={pageId}&"
                "c.mode={mode}&"
                "ds.{alias}.connector={connector}&"
                "ds.{alias}.spreadsheetId={spreadsheet_id}&"
                "ds.{alias}.worksheetId={worksheet_id}"
            )
            url = url.format(
                reportId=reportId,
                pageId=pageId,
                connector=connector,
                mode=mode,
                spreadsheet_id=spreadsheet_id,
                worksheet_id=worksheet_id,
                alias=alias,
            )
    else:
        pageId = pageId_list[0]
        url = (
            "https://lookerstudio.google.com/reporting/create?"
            "c.reportId={reportId}&"
            "c.pageId={pageId}&"
            "c.mode={mode}&"
            "ds.{alias}.connector={connector}&"
            "ds.{alias}.spreadsheetId={spreadsheet_id}&"
            "ds.{alias}.worksheetId={worksheet_id}"
        )

        url = url.format(
            reportId=reportId,
            pageId=pageId,
            connector=connector,
            mode=mode,
            spreadsheet_id=spreadsheet_id,
            worksheet_id=worksheet_id,
            alias=alias,
        )

    return url


# ---------- Shopify Export Helpers -----------


def get_country(country_mappinp: dict, country_code: str):
    """get the country name from the country code

    Args:
        country_mappinp (dict): given dict of country code and country name
        country_code (str): country code

    Returns:
        _type_: string
    """

    if country_code in country_mappinp.keys():
        return country_mappinp[country_code]
    else:
        return country_code


def get_shopify_export():
    """Get the shopify export file and pass it to process_file

    Returns:
        _type_: dataframe
    """
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
    )
    # file = input("Please select a file: ")
    # file = "data/dummy.xlsx"
    return process_file(file_path)


def process_file(file_path):
    if file_path.endswith(".csv"):
        try:
            df = pd.read_csv(file_path)
        except Exception as e:
            print(f"An error occurred: {e}")
            return None
    elif file_path.endswith(".xlsx"):
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            print(f"An error occurred: {e}")
            return None
    else:
        print("Please select a .csv or .xlsx file.")
        return get_shopify_export()  # Recursive call to retry

    df.columns = [
        "order_id",
        "product_name",
        "qty",
        "country",
        "order_created_at",
        "tracking_number",
    ]
    if "tracking_number" not in df.columns:
        print("Please select a file with a column named 'tracking_number'.")
        return get_shopify_export()  # Recursive call to retry

    df_main = df.dropna(axis=0, subset=["tracking_number"])
    print(f"There are {len(df)} orders with {len(df_main)} tracking numbers.")
    return df_main


def get_shipping_metrics(df_row: pd.Series):
    """get the shipping metrics from the dataframe row

    Args:
        df_row (pd.Series): shopify export with tracking infos

    Returns:
        _type_: tuple
    """
    # convert df_row string date into datetime date
    order_created_at = pd.to_datetime(df_row["order_created_at"], dayfirst=True)
    in_transit_at = pd.to_datetime(df_row["in_transit_at"])
    delivered_at = pd.to_datetime(df_row["delivered_at"])
    info_received_at = pd.to_datetime(df_row["info_received_at"])

    processing_time = (in_transit_at - order_created_at).days
    shipping_time = (delivered_at - in_transit_at).days
    total_time = processing_time + shipping_time

    order_created_at = (
        order_created_at.strftime("%Y-%m-%d") if pd.notnull(order_created_at) else None
    )
    info_received_at = (
        info_received_at.strftime("%Y-%m-%d") if pd.notnull(info_received_at) else None
    )
    in_transit_at = (
        in_transit_at.strftime("%Y-%m-%d") if pd.notnull(in_transit_at) else None
    )
    delivered_at = (
        delivered_at.strftime("%Y-%m-%d") if pd.notnull(delivered_at) else None
    )

    return (
        processing_time,
        shipping_time,
        total_time,
        order_created_at,
        info_received_at,
        in_transit_at,
        delivered_at,
    )
