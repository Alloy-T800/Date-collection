# Date-collection
# Alloy Composition and Performance Input Tool

This tool allows users to input alloy compositions and their respective performance properties through a graphical user interface (GUI). It supports automatic recognition of alloy elements, performance data input, phase properties input, and custom properties input. It also includes functionality for checking if an alloy already exists in the dataset and updating existing entries.

## Features

- **Automatic Alloy Composition Parsing**: Automatically recognizes and parses alloy compositions, including elements inside and outside parentheses.
- **Performance Property Input**: Allows users to input various performance properties of alloys.
- **Phase Properties Input**: Supports input of phase properties such as FCC, BCC, L12, etc.
- **Custom Properties Input**: Supports input of custom properties like eutectic structure and DOI.
- **Data Backup**: Automatically creates backups of the dataset before saving new entries.
- **Duplicate Alloy Checking**: Checks if an alloy already exists in the dataset and allows updating existing entries.

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/your-username/your-repo.git
    ```
2. Navigate to the project directory:
    ```sh
    cd your-repo
    ```
3. Install the required packages:
    ```sh
    pip install -r requirements.txt
    ```

## Usage

1. Ensure that you have a file named `data_set.xlsx` in the project directory. This file will be used to store and manage alloy data.
2. Run the main script:
    ```sh
    python your_script_name.py
    ```

## GUI Overview

### Main Window

- **Alloy Composition Input**: Enter the alloy composition in the input field and click "Submit" to process.
- **Result Area**: Displays the result of the alloy composition parsing and any duplicate alloy information.
- **Data Count**: Displays the number of entries in the dataset.

### New Alloy Input Window

- **Performance Properties**: Input fields for various performance properties such as T_YS, T_US, T_Elo, etc.
- **Phase Properties**: Radio buttons for selecting phase properties (FCC, BCC, L12, etc.).
- **Custom Properties**: Input fields for eutectic structure and DOI.
- **Submit Button**: Save the new alloy data to the dataset.

### Duplicate Alloy Update Window

- **Existing Alloy Information**: Displays existing alloy data that matches the new input.
- **Performance Properties**: Input fields for updating performance properties.
- **Phase Properties**: Radio buttons for updating phase properties.
- **Custom Properties**: Input fields for updating eutectic structure and DOI.
- **Save Button**: Save the updated alloy data to the dataset.

## Code Explanation

### Main Functions

- **backup_excel_file**: Creates a backup of the existing dataset before saving new entries.
- **parse_alloy_composition**: Parses the alloy composition input and calculates the proportions of each element.
- **process_property_input**: Validates and processes performance property inputs.
- **process_phase_properties**: Processes the phase properties input.
- **process_custom_properties**: Processes custom properties such as eutectic structure and DOI.
- **check_if_alloy_exists**: Checks if the alloy composition already exists in the dataset.
- **read_excel**: Reads the dataset from the Excel file.
- **write_excel**: Writes the updated dataset to the Excel file.
- **open_new_window_for_input**: Opens a new window for inputting performance data for new alloys.
- **save_updated_alloy_data**: Saves updated alloy data for existing alloys.
- **submit_performance_data**: Submits the performance data input for new alloys.
- **submit_alloy_data**: Handles the submission of alloy composition data from the main window.

## Contributing

1. Fork the repository.
2. Create a new branch:
    ```sh
    git checkout -b feature-branch
    ```
3. Make your changes and commit them:
    ```sh
    git commit -m 'Add some feature'
    ```
4. Push to the branch:
    ```sh
    git push origin feature-branch
    ```
5. Open a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Pandas](https://pandas.pydata.org/)
- [Tkinter](https://docs.python.org/3/library/tkinter.html)
- [NumPy](https://numpy.org/)
