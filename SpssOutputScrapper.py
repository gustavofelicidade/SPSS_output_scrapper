        # !/usr/bin/env python
        # encoding: utf-8

        class SpssOutputScrapper():

            def crawl_spss_output(wbs, spss_sheets, new_tables):
                value = ''
                for sheet_weight, spss_sheet in spss_sheets.items():  #
                    spss_sheets[sheet_weight] = wbs['new_output']['SPSS ' + sheet_weight]
                    for i, spss_row in enumerate(spss_sheet, 1):
                        if spss_row[0].value in table_names.values():  # Identify tables by unique titles
                            spss_table_key = [k for k, v in table_names.items() if v == spss_row[0].value][0]
                            for table_name, new_table in new_tables.items():
                                for row_key, new_row in new_table.items():
                                    for col_key, new_cell in new_row.items():
                                        new_table_key = new_tables[table_name][row_key][col_key]['key']
                                        weight = new_tables[table_name][row_key][col_key]['weight']
                                        if spss_table_key == new_table_key and sheet_weight == weight:
                                            row = getTableRow(row_key, i, spss_sheet)
                                            col = getTableColumn(row_key, i, spss_sheet)
                                            get_spss_value(row, col, spss_sheet, weight, new_tables, table_name,
                                                           row_key, col_key)
                return new_tables

            def getTableColumn(table_title, row, spss_sheet):

            # Crawl the first rows of the table to identify the labels
            # Stop to crawl when background is grey

            def getTableRow(rowType, row, spss_sheet):
                if rowType == 'Total':
                # Crawl the first column of the
                elif rowType == 'Count'
                # Identify valid values in second column & identify
                elif rowType == 'Prev'

            # Identify count and add 1 row (Prev/Perc always comes after count)

            def identifyTableEnd(cell):
        # Identify "total" in first column.
        # Or identify change i cell formattings (white vs grey background)
