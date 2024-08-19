import pyodbc
import datetime
import decimal

def clean_data(value, expected_type=None):
    if isinstance(value, str):
        value = value.strip()
        if value == '00:00:00':
            return None
        if value == '':
            return None
        if expected_type == 'int':
            return int(value) if value.isdigit() else None
        return value
    elif isinstance(value, datetime.datetime):
        if value == datetime.datetime(1899, 12, 30) or value.time() == datetime.time(0, 0):
            return None
        return value.strftime('%Y-%m-%d %H:%M:%S')
    elif isinstance(value, decimal.Decimal):
        return float(value)
    elif isinstance(value, bool):
        return 1 if value else 0
    elif expected_type == 'int':
        return int(value) if value is not None else None
    else:
        return value

def fetch_data_from_sql():
    dsn = "Orders"
    database = "FFOrders"
    try:
        sqlServer = pyodbc.connect(f"DSN={dsn};DATABASE={database};")
        sqlCursor = sqlServer.cursor()

        print(f"Connection to SQL Server: {dsn} and Database: {database} was successful!")

        fetchQuery_Heading = """
        SELECT AccountNo, AcknowledgementStatus, AmountPaid, AmountPaid_Converted, Architect, Archive,
               ArchivedCopyKeyID, ArchiveZip, BatchIndexOrder, BatchKeyID, BMGlassJob, BrickCavity,
               BuyUnglazedFrameGlass, BuyUnglazedFromSupplier, BuyPanelsFrom3rdParty, ChangeDate, ChangeTime,
               CheckedOutTo, CheckOrderCreated, Completed, Contact, Cost, CostLibrary, CounterUniqueKeyID,
               CreatedFrom, CreditAmount, CreditLimit, CreditNote, CreditNoteNumber, CurrencyCharacter, 
               CurrencyConversion, CustomerID, CustomerType, CustomerUseMarkupPerColour, CustomerUsePrimeAperture,
               CustomerMidrailHeight, CustomerTransomDrop, DateAmended, DateCompleted, DateConfirmed, DateCreated,
               DateDelivery, DateFitted, DateGlassDelivery, DateInProduction, DateOnwardDelivery, DateInvoice,
               DateLoaded, DateOrder, DatePaid, DatePaidSeveral, DatePanelDelivery, DateRoofDelivery, DateSawFileCreated,
               DateSurvey, DateToughGlassDelivery, DatePreOrderConverted, DatePreQuoteConverted, DeliveryAddress1,
               DeliveryAddress2, DeliveryAddress3, DeliveryAddress4, DeliveryCounty, DeliveryPostCode, DeliveryStatusID,
               DeliveryTeam, DeliveryTeamID, DepartmentID, DepositPaid, DepositPaid_Converted, Despatched, 
               Dessian_CurrencyConversion, DiscountFrame, DiscountFrame2, DiscountFrame3, DiscountFrame4, DiscountFrame5, 
               DiscountFrame6, DiscountFrame7, DiscountFrame8, DiscountFrame9, DiscountFrame10, Discount3rdPartyGlass, 
               Discount3rdPartyPanel, DiscountPartExtra, DiscountSATExtra, DiscountSATExtra2, DiscountSATExtra3, 
               DiscountSATExtra4, DiscountSATExtra5, DiscountSATExtra6, DiscountSATExtra7, DiscountSATExtra8, 
               DiscountSATExtra9, DiscountSATExtra10, Email, EmailLong, FaxNo, FensaNumber, FileName, Fitting, 
               FittingAdvancedID, FittingTeam, FittingTeamID, GlassAccountNo, GlassBatched, GlassSupplierID, Glazed, 
               Hidden, HousetypeCategoryID, Initials, InputBy, InputByID, InvoiceAddress1, InvoiceAddress2, 
               InvoiceAddress3, InvoiceAddress4, InvoiceCounty, InvoiceNumber, InvoicePostCode, InvoicePrinted, 
               InvoiceSettled, InvoiceSettledInformation, JobKeyID, JobType, MainColour
        FROM dbo.Heading
        WHERE SatelliteJobNumber = 'RB262';
        """

        sqlCursor.execute(fetchQuery_Heading)
        row = sqlCursor.fetchone()

        if row:
            print(f"Fetched row data: {row}")
            row = tuple(clean_data(value) for value in row)
            print(f"Cleaned row data: {row}")
            return row
        else:
            print("No data found for SatelliteJobNumber = 'RB262'.")
            return None

    except pyodbc.Error as ex:
        print(f"SQL Server connection or query failed: {ex}")
        return None

    finally:
        sqlCursor.close()
        sqlServer.close()
        print("SQL Connection Closed.")

def update_access_database(row):
    mdb = r"E:\BM\FuturaFrames\Orders\2024\Jun\Test\RB262.mdb"

    try:
        connStr = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={mdb};"
        )

        with pyodbc.connect(connStr) as accessConn:
            with accessConn.cursor() as accessCursor:
                print(f"Connection to Access Database: {mdb} was successful!")

                if row:
                    chunks = [
                        (24, """
                        UPDATE Heading
                        SET AccountNo = ?, AcknowledgementStatus = ?, AmountPaid = ?,
                            AmountPaid_Converted = ?, Architect = ?, Archive = ?, ArchivedCopyKeyID = ?, ArchiveZip = ?,
                            BatchIndexOrder = ?, BatchKeyID = ?, BMGlassJob = ?, BrickCavity = ?, BuyUnglazedFrameGlass = ?,
                            BuyUnglazedFromSupplier = ?, BuyPanelsFrom3rdParty = ?, ChangeDate = ?, ChangeTime = ?, CheckedOutTo = ?,
                            CheckOrderCreated = ?, Completed = ?, Contact = ?, Cost = ?, CostLibrary = ?, CounterUniqueKeyID = ?
                        WHERE SatelliteJobNumber = 'RB262';
                        """),
                        (24, """
                        UPDATE Heading
                        SET CreatedFrom = ?, CreditAmount = ?, CreditLimit = ?, CreditNote = ?,
                            CreditNoteNumber = ?, CurrencyCharacter = ?, CurrencyConversion = ?, CustomerID = ?,
                            CustomerType = ?, CustomerUseMarkupPerColour = ?, CustomerUsePrimeAperture = ?, CustomerMidrailHeight = ?,
                            CustomerTransomDrop = ?, DateAmended = ?, DateCompleted = ?, DateConfirmed = ?,
                            DateCreated = ?, DateDelivery = ?, DateFitted = ?, DateGlassDelivery = ?,
                            DateInProduction = ?, DateOnwardDelivery = ?, DateInvoice = ?, DateLoaded = ?
                        WHERE SatelliteJobNumber = 'RB262';
                        """),
                        (24, """
                        UPDATE Heading
                        SET DateOrder = ?, DatePaid = ?, DatePaidSeveral = ?, DatePanelDelivery = ?,
                            DateRoofDelivery = ?, DateSawFileCreated = ?, DateSurvey = ?, DateToughGlassDelivery = ?,
                            DatePreOrderConverted = ?, DatePreQuoteConverted = ?, DeliveryAddress1 = ?, DeliveryAddress2 = ?,
                            DeliveryAddress3 = ?, DeliveryAddress4 = ?, DeliveryCounty = ?, DeliveryPostCode = ?,
                            DeliveryStatusID = ?, DeliveryTeam = ?, DeliveryTeamID = ?, DepartmentID = ?,
                            DepositPaid = ?, DepositPaid_Converted = ?, Despatched = ?, Dessian_CurrencyConversion = ?
                        WHERE SatelliteJobNumber = 'RB262';
                        """),
                        (24, """
                        UPDATE Heading
                        SET DiscountFrame = ?, DiscountFrame2 = ?, DiscountFrame3 = ?, DiscountFrame4 = ?,
                            DiscountFrame5 = ?, DiscountFrame6 = ?, DiscountFrame7 = ?, DiscountFrame8 = ?,
                            DiscountFrame9 = ?, DiscountFrame10 = ?, Discount3rdPartyGlass = ?, Discount3rdPartyPanel = ?,
                            DiscountPartExtra = ?, DiscountSATExtra = ?, DiscountSATExtra2 = ?, DiscountSATExtra3 = ?,
                            DiscountSATExtra4 = ?, DiscountSATExtra5 = ?, DiscountSATExtra6 = ?, DiscountSATExtra7 = ?,
                            DiscountSATExtra8 = ?, DiscountSATExtra9 = ?, DiscountSATExtra10 = ?, Email = ?
                        WHERE SatelliteJobNumber = 'RB262';
                        """),
                        (30, """
                        UPDATE Heading
                        SET EmailLong = ?, FaxNo = ?, FensaNumber = ?, FileName = ?,
                            Fitting = ?, FittingAdvancedID = ?, FittingTeam = ?, FittingTeamID = ?,
                            GlassAccountNo = ?, GlassBatched = ?, GlassSupplierID = ?, Glazed = ?,
                            Hidden = ?, HousetypeCategoryID = ?, Initials = ?, InputBy = ?,
                            InputByID = ?, InvoiceAddress1 = ?, InvoiceAddress2 = ?, InvoiceAddress3 = ?,
                            InvoiceAddress4 = ?, InvoiceCounty = ?, InvoiceNumber = ?, InvoicePostCode = ?,
                            InvoicePrinted = ?, InvoiceSettled = ?, InvoiceSettledInformation = ?, JobKeyID = ?,
                            JobType = ?, MainColour = ?
                        WHERE SatelliteJobNumber = 'RB262';
                        """)
                    ]

                    for count, query in chunks:
                        print(f"Executing chunk with {count} placeholders.")
                        params = row[:count]
                        
                        for i, p in enumerate(params):
                            print(f"Parameter {i+1}: Value = {p}, Type = {type(p)}")
                        
                        print(f"SQL Query: {query}")
                        print(f"Number of parameters: {len(params)}")

                        try:
                            accessCursor.execute(query, *params)
                            accessConn.commit()
                            print(f"Chunk with {count} placeholders updated successfully.")
                        except pyodbc.Error as ex:
                            print(f"Chunk update failed: {ex}")
                            raise

                        row = row[count:]

                else:
                    print("No data to update in Access database.")

    except pyodbc.Error as ex:
        print(f"Access Database connection or update failed: {ex}")

    finally:
        try:
            accessCursor.close()
            accessConn.close()
        except NameError:
            pass
        print("Access Connection Closed.")

# Run the script
row_data = fetch_data_from_sql()
if row_data:
    update_access_database(row_data)
