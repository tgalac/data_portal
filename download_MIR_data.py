import io
import pandas as pd
import requests


def get_ecb_mir_croatia():
    # 1. Define the ECB SDMX API endpoint
    # Base format: https://data-api.ecb.europa.eu/v1/data/{flowRef}/{key}
    base_url = "https://data-api.ecb.europa.eu/service/data/MIR/"

    # 2. Construct the wildcard key according to your parameters
    # Dimensions: Freq(M).Area(HR).Type.Instrument.Maturity.Amount.Coverage.Sector.Currency(EUR).Suffix(N)
    # Leaving intermediate dimensions blank acts as a wildcard to pull ALL variations
    series_key = "M.HR.......EUR.N"

    # 3. Request data specifically in CSV format for easy Pandas integration
    headers = {"Accept": "text/csv"}

    print(f"Requesting data from ECB API for key: MIR.{series_key}...")
    response = requests.get(base_url + series_key, headers=headers)

    # 4. Check if request was successful
    if response.status_code == 200:
        print("Success! Loading into a Pandas DataFrame...")

        # Read the raw text CSV into Pandas
        df = pd.read_csv(io.StringIO(response.text))

        # Sort values nicely by Series Code and Time Period
        if "KEY" in df.columns and "TIME_PERIOD" in df.columns:
            df = df.sort_values(by=["KEY", "TIME_PERIOD"]).reset_index(
                drop=True
            )

        return df
    else:
        print(f"Failed to fetch data. HTTP Status Code: {response.status_code}")
        print(response.text)
        return None


# Run the script
if __name__ == "__main__":
    mir_hr_df = get_ecb_mir_croatia()

    if mir_hr_df is not None:
      
        # Save it locally to a CSV file
        mir_hr_df.to_csv("MIR_podaci_HR.csv", index=False)
