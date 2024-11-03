import pandas as pd
import random
from datetime import datetime, timedelta


class InventorySalesDataGenerator:
    def __init__(self, num_inventory_rows=200, num_sales_rows=25000):
        """
        Initialize the data generator with specified number of inventory and sales rows.

        Args:
            num_inventory_rows (int): Number of inventory records to generate.
            num_sales_rows (int): Number of sales records to generate.
        """
        self.num_inventory_rows = num_inventory_rows
        self.num_sales_rows = num_sales_rows
        self.categories = ["Odzież", "Obuwie", "Akcesoria"]
        self.cities = ["Warszawa", "Wrocław", "Gdańsk", "Poznań", "Kraków"]
        self.start_date = datetime(2024, 10, 1)
        self.end_date = datetime(2024, 10, 31)

    def generate_inventory_data(self):
        """
        Generate inventory data with product details and stock levels.

        Returns:
            pd.DataFrame: DataFrame containing inventory data.
        """
        inventory_data = {
            "id produktu": [f"P{str(i).zfill(4)}" for i in range(1, self.num_inventory_rows + 1)],
            "nazwa produktu": [f"Produkt_{i}" for i in range(1, self.num_inventory_rows + 1)],
            "kategoria": [random.choice(self.categories) for _ in range(self.num_inventory_rows)],
            "cena zakupu": [round(random.uniform(10, 500), 2) for _ in range(self.num_inventory_rows)],
            "cena sprzedaży": [round(random.uniform(10, 700), 2) for _ in range(self.num_inventory_rows)],
            **{f"stan zapasu Sklep {city}": [random.randint(0, 50) for _ in range(self.num_inventory_rows)] for city in self.cities},
            "stan zapasu Magazyn": [random.randint(0, 200) for _ in range(self.num_inventory_rows)],
        }
        return pd.DataFrame(inventory_data)

    def generate_sales_data(self):
        """
        Generate sales data with product sales records.

        Returns:
            pd.DataFrame: DataFrame containing sales data.
        """
        sales_data = {
            "id produktu": [f"P{str(random.randint(1, self.num_inventory_rows)).zfill(4)}" for _ in range(self.num_sales_rows)],
            "data": [
                (self.start_date + timedelta(days=random.randint(0, (self.end_date - self.start_date).days))).strftime("%d.%m.%Y")
                for _ in range(self.num_sales_rows)],
            "ilosc": [random.randint(1, 10) for _ in range(self.num_sales_rows)],
            "sklep": [random.choice(self.cities) for _ in range(self.num_sales_rows)],
        }
        return pd.DataFrame(sales_data)

    def create_excel_file(self, filename="output.xlsx"):
        """
        Create an Excel file with generated inventory and sales data.

        Args:
            filename (str): The name of the output Excel file.
        """
        inventory_data = self.generate_inventory_data()
        sales_data = self.generate_sales_data()

        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            inventory_data.to_excel(writer, index=False, sheet_name="Stan zapasów 2024.10.31")
            sales_data.to_excel(writer, index=False, sheet_name="Sprzedaż za 2024.10")

        print(f"{filename} created successfully.")


if __name__ == "__main__":
    data_generator = InventorySalesDataGenerator()
    data_generator.create_excel_file("inventory_and_sales_data.xlsx")
