from create_databases import ReadandWrite

def main():
    input_files = [
            '../Read/Ex_input_file_02.04.2021.xlsx',
            '../Read/Ex_mapping_file.xlsx'
    ]
    read_and_write = ReadandWrite()
#     read_and_write.excel_to_db(input_files)
    read_and_write.run_query("SELECT * FROM 'index data'")


if __name__ == "__main__":
        main();


