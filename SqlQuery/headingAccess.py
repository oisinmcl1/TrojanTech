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

def fetch_data_from_sql(satellite_job_number):
    dsn = "Orders"
    database = "FFOrders"
    try:
        # print(f"Attempting to connect to SQL Server: DSN={dsn}, Database={database}")
        sqlServer = pyodbc.connect(f"DSN={dsn};DATABASE={database};")
        sqlCursor = sqlServer.cursor()

        # print(f"Connection to SQL Server was successful!")
        # print(f"Executing SQL query...")

        fetchQuery_Heading = f"""
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
               InvoiceSettled, InvoiceSettledInformation, JobKeyID, JobNumber, JobType, MainColour,
               MasterJobKeyID, Mobile, Name, NettDeliveryCharge, NettDeliveryCharge_Converted, Notes, OmitGlazingMethodGlassFactory,
               OmitGlazingMethodGlassSite, OmitGlazingMethodUnglazedSpecific, OmitGlazingMethodUnglazedGeneric, OmitGlazingMethodPanelFactory,
               OmitGlazingMethodPanelSite, OmitGlazingMethodPanelUnglazedSpecific, OmitGlazingMethodPanelUnglazedGeneric,
               OmitStock, OnHold, OrderbookID, OrderStatus, PanelSupplierID, PaymentMethod, PaymentMethodID, PhoneNo, Posted,
               PreOrderPrinted, PreQuotePrinted, PriceGrids, Projectname, QuantityFabricationUnits, QuantityFabricationUnits_Fabrication,
               QuantityFrames, QuantityGlass, QuantityGlassAnnealed, QuantityGlassLaminated, QuantityGlassObscure, QuantityGlassToughened,
               QuantityPanels, QuantityPanelsFlat, QuantityPanelsMoulded, QuantityRoofs, QuantityRoofPacks, QuantityUnglazed,
               QuoteNumber, QuotePrinted, ReceivedOrder, Reference, Remake, RemakeCount, RemakePro, RemakeTypeID,
               RoofwrightRoofCost, RoofwrightRoofPrice, RoofwrightBaseCost, RoofwrightBasePrice, RouteKeyID,
               Salesman, SalesmanCommission, SalesmanID, Salutation, SatelliteCustomerID, SatelliteJobKeyID,
               SatelliteJobNumber, SatelliteName, SatelliteUploadedDate, SellingPriceExTax, SellingPriceExTax_Converted,
               SellingPriceIncTax, SellingPriceIncTax_Converted, SettlementDiscount, Status, StockExtraCostPlus, 
               SurchargePercent, SurchargeValue, Surveyor, SurveyorID, TaxAmount, TaxAmount_Converted, TaxCode,
               TaxRate, Toughened, VersionCreated, VersionAmended, Waypoint, WER, WERid, Urgent, UValue
        FROM dbo.Heading
        WHERE JobNumber = '{satellite_job_number}';
        """

        sqlCursor.execute(fetchQuery_Heading)
        row = sqlCursor.fetchone()

        if row:
            # print("Fetched row data from SQL Server:")
            # for idx, value in enumerate(row):
            #     print(f"Column {idx+1}: {value} (Type: {type(value)})")
                
            row = tuple(clean_data(value) for value in row)

            # print("\nCleaned data to be updated in Access Database:")
            # for idx, value in enumerate(row):
            #     print(f"Column {idx+1}: {value} (Type: {type(value)})")
            
            return row
        else:
            print(f"No data found for SatelliteJobNumber = '{satellite_job_number}'.")
            return None

    except pyodbc.Error as ex:
        print(f"SQL Server connection or query failed: {ex}")
        return None

    finally:
        sqlCursor.close()
        sqlServer.close()
        # print("SQL Connection Closed.")

def update_access_database(row, job_number):
    mdb = f"C:\\temp\\attachments\\{job_number}.mdb"

    try:
        # print(f"Attempting to connect to Access Database: {mdb}")
        connStr = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={mdb};"
        )

        with pyodbc.connect(connStr) as accessConn:
            with accessConn.cursor() as accessCursor:
                # print(f"Connection to Access Database: {mdb} was successful!")

                if row:
                    # Update queries broken into chunks
                    update_queries = [
                        """
                        UPDATE Heading SET 
                            AccountNo = ?, AcknowledgementStatus = ?, AmountPaid = ?, AmountPaid_Converted = ?, 
                            Architect = ?, Archive = ?, ArchivedCopyKeyID = ?, ArchiveZip = ?, 
                            BatchIndexOrder = ?, BatchKeyID = ?, BMGlassJob = ?, BrickCavity = ?, 
                            BuyUnglazedFrameGlass = ?, BuyUnglazedFromSupplier = ?, BuyPanelsFrom3rdParty = ?, ChangeDate = ?, 
                            ChangeTime = ?, CheckedOutTo = ?, CheckOrderCreated = ?, Completed = ?, 
                            Contact = ?, Cost = ?, CostLibrary = ?, CounterUniqueKeyID = ?, 
                            CreatedFrom = ?, CreditAmount = ?, CreditLimit = ?, CreditNote = ?, 
                            CreditNoteNumber = ?, CurrencyCharacter = ?, CurrencyConversion = ?, CustomerID = ?
                        WHERE JobNumber = ?;
                        """,
                        """
                        UPDATE Heading SET 
                            CustomerType = ?, CustomerUseMarkupPerColour = ?, CustomerUsePrimeAperture = ?, CustomerMidrailHeight = ?, 
                            CustomerTransomDrop = ?, DateAmended = ?, DateCompleted = ?, DateConfirmed = ?, 
                            DateCreated = ?, DateDelivery = ?, DateFitted = ?, DateGlassDelivery = ?, 
                            DateInProduction = ?, DateOnwardDelivery = ?, DateInvoice = ?, DateLoaded = ?, 
                            DateOrder = ?, DatePaid = ?, DatePaidSeveral = ?, DatePanelDelivery = ?, 
                            DateRoofDelivery = ?, DateSawFileCreated = ?, DateSurvey = ?, DateToughGlassDelivery = ?, 
                            DatePreOrderConverted = ?, DatePreQuoteConverted = ?, DeliveryAddress1 = ?, DeliveryAddress2 = ?, 
                            DeliveryAddress3 = ?, DeliveryAddress4 = ?, DeliveryCounty = ?, DeliveryPostCode = ?
                        WHERE JobNumber = ?;
                        """,
                        """
                        UPDATE Heading SET 
                            DeliveryStatusID = ?, DeliveryTeam = ?, DeliveryTeamID = ?, DepartmentID = ?, 
                            DepositPaid = ?, DepositPaid_Converted = ?, Despatched = ?, Dessian_CurrencyConversion = ?, 
                            DiscountFrame = ?, DiscountFrame2 = ?, DiscountFrame3 = ?, DiscountFrame4 = ?, 
                            DiscountFrame5 = ?, DiscountFrame6 = ?, DiscountFrame7 = ?, DiscountFrame8 = ?, 
                            DiscountFrame9 = ?, DiscountFrame10 = ?, Discount3rdPartyGlass = ?, Discount3rdPartyPanel = ?, 
                            DiscountPartExtra = ?, DiscountSATExtra = ?, DiscountSATExtra2 = ?, DiscountSATExtra3 = ?, 
                            DiscountSATExtra4 = ?, DiscountSATExtra5 = ?, DiscountSATExtra6 = ?, DiscountSATExtra7 = ?, 
                            DiscountSATExtra8 = ?, DiscountSATExtra9 = ?, DiscountSATExtra10 = ?
                        WHERE JobNumber = ?;
                        """,
                        """
                        UPDATE Heading SET 
                            Email = ?, EmailLong = ?, FaxNo = ?, FensaNumber = ?, 
                            FileName = ?, Fitting = ?, FittingAdvancedID = ?, 
                            FittingTeam = ?, FittingTeamID = ?, GlassAccountNo = ?, GlassBatched = ?, 
                            GlassSupplierID = ?, Glazed = ?, Hidden = ?, HousetypeCategoryID = ?, 
                            Initials = ?, InputBy = ?, InputByID = ?, InvoiceAddress1 = ?, 
                            InvoiceAddress2 = ?, InvoiceAddress3 = ?, InvoiceAddress4 = ?, InvoiceCounty = ?, 
                            InvoiceNumber = ?, InvoicePostCode = ?, InvoicePrinted = ?, InvoiceSettled = ?, 
                            InvoiceSettledInformation = ?, JobKeyID = ?, JobNumber = ?, JobType = ?
                        WHERE JobNumber = ?;
                        """,
                        """
                        UPDATE Heading SET 
                            MainColour = ?, MasterJobKeyID = ?, Mobile = ?, 
                            Name = ?, NettDeliveryCharge = ?, NettDeliveryCharge_Converted = ?, Notes = ?, 
                            OmitGlazingMethodGlassFactory = ?, OmitGlazingMethodGlassSite = ?, OmitGlazingMethodUnglazedSpecific = ?, 
                            OmitGlazingMethodUnglazedGeneric = ?, OmitGlazingMethodPanelFactory = ?, OmitGlazingMethodPanelSite = ?, 
                            OmitGlazingMethodPanelUnglazedSpecific = ?, OmitGlazingMethodPanelUnglazedGeneric = ?, OmitStock = ?, 
                            OnHold = ?, OrderbookID = ?, OrderStatus = ?, PanelSupplierID = ?, 
                            PaymentMethod = ?, PaymentMethodID = ?, PhoneNo = ?, Posted = ?, 
                            PreOrderPrinted = ?, PreQuotePrinted = ?, PriceGrids = ?, Projectname = ?, 
                            QuantityFabricationUnits = ?, QuantityFabricationUnits_Fabrication = ?, QuantityFrames = ?, QuantityGlass = ?, 
                            QuantityGlassAnnealed = ?, QuantityGlassLaminated = ?, QuantityGlassObscure = ?, QuantityGlassToughened = ?, 
                            QuantityPanels = ?, QuantityPanelsFlat = ?, QuantityPanelsMoulded = ?, QuantityRoofs = ?, 
                            QuantityRoofPacks = ?, QuantityUnglazed = ?, QuoteNumber = ?, QuotePrinted = ?, 
                            ReceivedOrder = ?, Reference = ?, Remake = ?, RemakeCount = ?, 
                            RemakePro = ?, RemakeTypeID = ?, RoofwrightRoofCost = ?, RoofwrightRoofPrice = ?, 
                            RoofwrightBaseCost = ?, RoofwrightBasePrice = ?, RouteKeyID = ?, Salesman = ?, 
                            SalesmanCommission = ?, SalesmanID = ?, Salutation = ?, SatelliteCustomerID = ?, 
                            SatelliteJobKeyID = ?, SatelliteJobNumber = ?, SatelliteName = ?, SatelliteUploadedDate = ?, 
                            SellingPriceExTax = ?, SellingPriceExTax_Converted = ?, SellingPriceIncTax = ?, SellingPriceIncTax_Converted = ?, 
                            SettlementDiscount = ?, Status = ?, StockExtraCostPlus = ?, SurchargePercent = ?, 
                            SurchargeValue = ?, Surveyor = ?, SurveyorID = ?, TaxAmount = ?, 
                            TaxAmount_Converted = ?, TaxCode = ?, TaxRate = ?, Toughened = ?, 
                            VersionCreated = ?, VersionAmended = ?, Waypoint = ?, WER = ?, 
                            WERid = ?, Urgent = ?, UValue = ?
                        WHERE JobNumber = ?;
                        """
                    ]

                    # Adjust chunk sizes according to the number of placeholders
                    chunk_sizes = [32, 32, 31, 31, 88]  # Adjusted to match the query placeholders

                    # Execute each chunk
                    start_idx = 0
                    for idx, (query, chunk_size) in enumerate(zip(update_queries, chunk_sizes)):
                        end_idx = start_idx + chunk_size
                        params = row[start_idx:end_idx] + (job_number,)

                        # Debugging outputs
                        # print(f"Executing update query {idx+1}...")
                        # print(f"Query: {query}")
                        # print(f"Number of placeholders: {query.count('?')}")
                        # print(f"Params{idx+1}: {params}")
                        # print(f"Number of parameters provided: {len(params)}")

                        assert len(params) == query.count('?'), f"Mismatch in query {idx+1} between placeholders and parameters."

                        accessCursor.execute(query, params)
                        start_idx += chunk_size  # Move to the next set of parameters

                    # Commit all changes
                    accessConn.commit()
                    # print("All updates committed successfully.")

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
        # print("Access Connection Closed.")

def main(satellite_job_number):
    # This function encapsulates the main logic of the script
    
    # Fetch data from SQL Server
    row_data = fetch_data_from_sql(satellite_job_number)
    
    if row_data:
        # Update Access Database with the fetched data
        update_access_database(row_data, satellite_job_number)
        print(f"Script ran successfully for heading job: {satellite_job_number}")
    else:
        print(f"No data was returned from the SQL query for job {satellite_job_number}.")

if __name__ == "__main__":
    pass
