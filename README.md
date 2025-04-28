# SunA - Solar Energy Calculator

SunA is a Python-based application that calculates the potential clear-sky solar energy generation for a 1m² horizontal solar panel at specified locations. It provides monthly energy production estimates in kWh over a 10-year period, using the `pvlib` library for solar position and irradiance calculations. The results are exported to an Excel file for easy analysis.

## Features
- Calculates monthly clear-sky solar energy production (kWh/month) for a 1m² panel.
- Supports multiple locations with customizable latitude, longitude, and timezone.
- User-friendly GUI built with `tkinter`.
- Allows specification of panel efficiency and starting year.
- Exports results to an Excel file with a formatted sheet name.

## Prerequisites
- Python 3.6 or higher
- Required Python packages (listed in `requirements.txt`)

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/mnsoylemez/SunA.git
   cd SunA
   ```
2. Create a virtual environment (optional but recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the application:
   ```bash
   python solar_energy_calculator.py
   ```

## Usage
1. Launch the application by running `solar_energy_calculator.py`.
2. Enter the following details in the GUI:
   - **Location 1 and 2**: Specify the name, latitude (°), longitude (°), and timezone (select from the dropdown, e.g., GMT+3).
   - **Start Year**: Enter the starting year for the 10-year calculation (e.g., 2024).
   - **Panel Efficiency (%)**: Enter the solar panel efficiency as a percentage (e.g., 20 for 20%).
3. Click **"Hesapla ve Aylık Enerjiyi (kWh) Dışa Aktar"** to calculate and export the results.
4. Choose a file path to save the Excel file when prompted.
5. The results will be saved in an Excel file with monthly energy production (kWh/month) for each location.

## Output
- The Excel file contains a sheet named `Aylik_kWh_Xverim` (where `X` is the panel efficiency percentage).
- Columns are labeled as `[Location Name] Açık Hava Enerji (kWh/ay, %X verim)`.
- Rows are indexed by year and month (e.g., `2024-01`).

## Notes
- The calculation assumes clear-sky conditions and does not account for weather variations, shading, or panel tilt.
- The application is designed for a 1m² horizontal panel.
- Timezones must be selected from the provided list (e.g., GMT+3 corresponds to `Etc/GMT-3`).

## Dependencies
See `requirements.txt` for the full list of dependencies. Key libraries include:
- `tkinter` (for the GUI)
- `pandas` (for data handling)
- `numpy` (for numerical operations)
- `pvlib` (for solar calculations)

## License
See the `LICENSE` file for details.

## Contributing
Contributions are welcome! Please open an issue or submit a pull request on GitHub.

## Contact
For questions or feedback, please contact soylemeznurhan@gmail.com.
