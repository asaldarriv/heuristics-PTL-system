# Random Method

This project implements a random method, specifically an evolutionary method (1+1) that uses an initial solution provided by the nearest neighbor method (constructive method) and then improves it by making small mutations.

## Description

- **Deterministic Method (Nearest Neighbor)**: orders are sorted in descending order by the number of SKUs. Once the order is determined, each order is assigned to the nearest exit.
- **Random Method (Evolutionaty Method (1+1))**: uses an initial solution provided by the nearest neighbor method (constructive method) and then improves it by making small mutations.

## Running the Algorithm

### Requirements

- Python 3.11.3

### Instructions for Windows

1. Install `virtualenv`:
    ```sh
    pip install virtualenv
    ```

2. Create a virtual environment in the project's root directory:
    ```sh
    virtualenv <virtual_environment_name>
    ```

3. Activate the virtual environment:
    ```sh
    source <virtual_environment_name>/Scripts/activate
    ```

4. Install the dependencies:
    ```sh
    pip install -r requirements.txt
    ```

5. Run the algorithm:
    ```sh
    python random_method/main.py
    ```

    ## Results

    After running the algorithm, the results are stored in several folders:

    - **solutions**: contains an Excel file for each instance/method.
    - **reports**: contains an Excel file comparing the results of different methods and instances, as well as an image comparing the workload distribution across zones.
    - **bar_images**: contains bar charts showing the workload balance for each instance/method.

    These files provide a detailed analysis of the performance and efficiency of the different methods applied to the instances.
