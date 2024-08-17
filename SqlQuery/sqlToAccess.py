import pyodbc
import datetime
import decimal

def fetch_data_from_sql():
    """
    Connect to the SQL Server and fetch the row data from dbo.Heading.
    """
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
               CreatedFrom, CustomerID, CustomerType, CustomerUseMarkupPerColour, CustomerUsePrimeAperture,
               CustomerMidrailHeight, CustomerTransomDrop, DateAmended, DateCompleted, DateConfirmed, DateCreated,
               DateDelivery, DateFitted, DateGlassDelivery, DateInProduction, DateOnwardDelivery, DateInvoice,
               DateLoaded, DateOrder, DatePaid, DatePaidSeveral, DatePanelDelivery, DateRoofDelivery, DateSawFileCreated,
               DateSurvey, DateToughGlassDelivery, DatePreOrderConverted, DatePreQuoteConverted, Despatched,
               Dessian_CurrencyConversion, Email, EmailLong, FaxNo, FensaNumber, FileName, Fitting, FittingAdvancedID,
               FittingTeam, FittingTeamID, GlassAccountNo, GlassBatched, GlassSupplierID, Glazed, Hidden, HousetypeCategoryID,
               Initials, InputBy, InputByID, InvoiceAddress1, InvoiceAddress2, InvoiceAddress3, InvoiceAddress4,
               InvoiceCounty, InvoiceNumber, InvoicePostCode, InvoicePrinted, InvoiceSettled, InvoiceSettledInformation,
               JobKeyID, JobNumber, JobType, MainColour, MasterJobKeyID, Mobile, Name, NettDeliveryCharge,
               NettDeliveryCharge_Converted, Notes, PanelSupplierID, PaymentMethod, PaymentMethodID, PhoneNo, Posted,
               PreOrderPrinted, PreQuotePrinted, PriceGrids, Projectname, QuantityFabricationUnits,
               QuantityFabricationUnits_Fabrication, QuantityFrames, QuantityGlass, QuantityGlassAnnealed,
               QuantityGlassLaminated, QuantityGlassObscure, QuantityGlassToughened, QuantityPanels, QuantityPanelsFlat,
               QuantityPanelsMoulded, QuantityRoofs, QuantityRoofPacks, QuantityUnglazed, QuoteNumber, QuotePrinted,
               ReceivedOrder, Reference, RouteKeyID, Salesman, SalesmanCommission, SalesmanID, Salutation,
               SatelliteCustomerID, SatelliteJobKeyID, SatelliteJobNumber, SatelliteName, SatelliteUploadedDate,
               SellingPriceExTax, SellingPriceExTax_Converted, SellingPriceIncTax, SellingPriceIncTax_Converted, TaxAmount,
               TaxAmount_Converted, TaxCode, TaxRate, Toughened, VersionCreated, VersionAmended
        FROM dbo.Heading
        WHERE SatelliteJobNumber = 'RB262';
        """

        sqlCursor.execute(fetchQuery_Heading)
        row = sqlCursor.fetchone()

        if row:
            print(f"Fetched {len(row)} columns from the SQL Server database.")
            # Clean up the data
            row = tuple(
                None if value is None or (isinstance(value, str) and not value.strip()) else 
                None if isinstance(value, datetime.datetime) and value == datetime.datetime(1899, 12, 30) else 
                str(value) if isinstance(value, datetime.datetime) else 
                float(value) if isinstance(value, decimal.Decimal) else 
                1 if value is True else 0 if value is False else 
                value 
                for value in row
            )
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
    """
    Connect to the Access Database and update the Heading table.
    """
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
                        WHERE JobNumber = 'RB262';
                        """),
                        (24, """
                        UPDATE Heading
                        SET CreatedFrom = ?, CustomerID = ?, CustomerType = ?, CustomerUseMarkupPerColour = ?,
                            CustomerUsePrimeAperture = ?, CustomerMidrailHeight = ?, CustomerTransomDrop = ?, DateAmended = ?,
                            DateCompleted = ?, DateConfirmed = ?, DateCreated = ?, DateDelivery = ?, DateFitted = ?, DateGlassDelivery = ?,
                            DateInProduction = ?, DateOnwardDelivery = ?, DateInvoice = ?, DateLoaded = ?, DateOrder = ?, DatePaid = ?,
                            DatePaidSeveral = ?, DatePanelDelivery = ?, DateRoofDelivery = ?, DateSawFileCreated = ?
                        WHERE JobNumber = 'RB262';
                        """),
                        (23, """
                        UPDATE Heading
                        SET DateToughGlassDelivery = ?, DatePreOrderConverted = ?, DatePreQuoteConverted = ?, Despatched = ?,
                            Dessian_CurrencyConversion = ?, Email = ?, EmailLong = ?, FaxNo = ?, FensaNumber = ?, FileName = ?, Fitting = ?,
                            FittingAdvancedID = ?, FittingTeam = ?, FittingTeamID = ?, GlassAccountNo = ?, GlassBatched = ?,
                            GlassSupplierID = ?, Glazed = ?, Hidden = ?, HousetypeCategoryID = ?, Initials = ?, InputBy = ?, InputByID = ?
                        WHERE JobNumber = 'RB262';
                        """),
                        (24, """
                        UPDATE Heading
                        SET InvoiceAddress1 = ?, InvoiceAddress2 = ?, InvoiceAddress3 = ?, InvoiceAddress4 = ?, InvoiceCounty = ?,
                            InvoiceNumber = ?, InvoicePostCode = ?, InvoicePrinted = ?, InvoiceSettled = ?, InvoiceSettledInformation = ?,
                            JobKeyID = ?, JobNumber = ?, JobType = ?, MainColour = ?, MasterJobKeyID = ?, Mobile = ?, Name = ?,
                            NettDeliveryCharge = ?, NettDeliveryCharge_Converted = ?, Notes = ?, PanelSupplierID = ?, PaymentMethod = ?,
                            PaymentMethodID = ?, PhoneNo = ?, Posted = ?, PreOrderPrinted = ?, PreQuotePrinted = ?, PriceGrids = ?,
                            Projectname = ?, QuantityFabricationUnits = ?, QuantityFabricationUnits_Fabrication = ?, QuantityFrames = ?
                        WHERE JobNumber = 'RB262';
                        """),
                        (25, """
                        UPDATE Heading
                        SET QuantityGlass = ?, QuantityGlassAnnealed = ?, QuantityGlassLaminated = ?, QuantityGlassObscure = ?,
                            QuantityGlassToughened = ?, QuantityPanels = ?, QuantityPanelsFlat = ?, QuantityPanelsMoulded = ?,
                            QuantityRoofs = ?, QuantityRoofPacks = ?, QuantityUnglazed = ?, QuoteNumber = ?, QuotePrinted = ?,
                            ReceivedOrder = ?, Reference = ?, RouteKeyID = ?, Salesman = ?, SalesmanCommission = ?, SalesmanID = ?,
                            Salutation = ?, SatelliteCustomerID = ?, SatelliteJobKeyID = ?, SatelliteJobNumber = ?,
                            SatelliteName = ?, SatelliteUploadedDate = ?, SellingPriceExTax = ?, SellingPriceExTax_Converted = ?,
                            SellingPriceIncTax = ?, SellingPriceIncTax_Converted = ?, TaxAmount = ?, TaxAmount_Converted = ?,
                            TaxCode = ?, TaxRate = ?, Toughened = ?, VersionCreated = ?, VersionAmended = ?
                        WHERE JobNumber = 'RB262';
                        """)
                    ]

                    for count, query in chunks:
                        print(f"Executing chunk with {count} placeholders.")
                        params = row[:count]
                        
                        # Convert boolean and datetime to appropriate types
                        params = tuple(
                            1 if p is True else 0 if p is False else
                            None if isinstance(p, str) and p == '' else  # Convert empty strings to None
                            None if isinstance(p, datetime.datetime) and p == datetime.datetime(1899, 12, 30) else  # Handle placeholder date
                            str(p) if isinstance(p, datetime.datetime) else  # Convert datetime to string if not placeholder
                            float(p) if isinstance(p, decimal.Decimal) else  # Convert decimal to float
                            p for p in params
                        )
                        
                        # Debug: Print each parameter type and value
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

                        # Update row to exclude already used parameters
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
