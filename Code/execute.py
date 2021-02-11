from create_databases import ReadandWrite

def main():
    input_files = [
            '../Read/Ex_input_file_02.04.2021.xlsx',
            '../Read/Ex_mapping_file.xlsx'
    ]
    generated_files = [
        'DAV Proforma Acc Analy.xlsx',
        'Portfolio Account NAV.xlsx'
    ]
    read_and_write = ReadandWrite()
    # read_and_write.excel_to_db(input_files)
    # read_and_write.run_query("SELECT * FROM 'index data'")
    # print(read_and_write.extract_timestamp_from_file(input_files[0]))
    # read_and_write.generate_portfolio_valuation_file1(input_files, "DAV Proforma Acc Analy.xlsx")
    # read_and_write.generate_portfolio_valuation_file2(input_files[0], "Portfolio Account NAV.xlsx")
    # read_and_write.excel_to_db(generated_files)
    read_and_write.convert_xlsx_to_xml(generated_files)

if __name__ == "__main__":
        main();


