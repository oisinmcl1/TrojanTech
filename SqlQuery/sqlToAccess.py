import pyodbc

"""
CONNECT TO SQL SERVER
FETCH DBO.HEADING
"""

dsn = "Orders"
database = "FFOrders"

# Try to connect to the SQL Server database and create a cursor
try:
    sqlServer = pyodbc.connect(f"DSN={dsn};DATABASE={database};")
    sqlCursor = sqlServer.cursor()

    print(f"Connection to SQL Server: {dsn} and Database: {database} was successful!")

    # Fetch rows from dbo.Heading where SatelliteJobNumber = 'RB262'
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

    """
    # Verify fetched data
    if row:
        print("Fetched data:")
        print(row)
    else:
        print("No data found for SatelliteJobNumber = 'RB262'.")
    """

except pyodbc.Error as ex:
    print(f"SQL Server connection or query failed: {ex}")
    row = None

finally:
    sqlCursor.close()
    sqlServer.close()
    print("SQL Connection Closed.")


"""
CONNECT TO ACCESS DATABASE
UPDATE HEADING
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
                # First chunk of fields
                updateQuery_Heading_1 = """
                UPDATE Heading
                SET AccountNo = ?, AcknowledgementStatus = ?, AmountPaid = ?,
                    AmountPaid_Converted = ?, Architect = ?, Archive = ?, ArchivedCopyKeyID = ?, ArchiveZip = ?,
                    BatchIndexOrder = ?, BatchKeyID = ?, BMGlassJob = ?, BrickCavity = ?, BuyUnglazedFrameGlass = ?,
                    BuyUnglazedFromSupplier = ?, BuyPanelsFrom3rdParty = ?, ChangeDate = ?, ChangeTime = ?, CheckedOutTo = ?,
                    CheckOrderCreated = ?, Completed = ?, Contact = ?, Cost = ?, CostLibrary = ?, CounterUniqueKeyID = ?
                WHERE JobNumber = 'RB262';
                """

                # Execute the first chunk
                accessCursor.execute(updateQuery_Heading_1, *row[:24])
                accessConn.commit()

                # Second chunk of fields
                updateQuery_Heading_2 = """
                UPDATE Heading
                SET CreatedFrom = ?, CustomerID = ?, CustomerType = ?, CustomerUseMarkupPerColour = ?,
                    CustomerUsePrimeAperture = ?, CustomerMidrailHeight = ?, CustomerTransomDrop = ?, DateAmended = ?,
                    DateCompleted = ?, DateConfirmed = ?, DateCreated = ?, DateDelivery = ?, DateFitted = ?, DateGlassDelivery = ?,
                    DateInProduction = ?, DateOnwardDelivery = ?, DateInvoice = ?, DateLoaded = ?, DateOrder = ?, DatePaid = ?,
                    DatePaidSeveral = ?, DatePanelDelivery = ?, DateRoofDelivery = ?, DateSawFileCreated = ?, DateSurvey = ?
                WHERE JobNumber = 'RB262';
                """

                # Execute the second chunk
                accessCursor.execute(updateQuery_Heading_2, *row[24:48])
                accessConn.commit()

                # Third chunk of fields
                updateQuery_Heading_3 = """
                UPDATE Heading
                SET DateToughGlassDelivery = ?, DatePreOrderConverted = ?, DatePreQuoteConverted = ?, Despatched = ?,
                    Dessian_CurrencyConversion = ?, Email = ?, EmailLong = ?, FaxNo = ?, FensaNumber = ?, FileName = ?, Fitting = ?,
                    FittingAdvancedID = ?, FittingTeam = ?, FittingTeamID = ?, GlassAccountNo = ?, GlassBatched = ?,
                    GlassSupplierID = ?, Glazed = ?, Hidden = ?, HousetypeCategoryID = ?, Initials = ?, InputBy = ?, InputByID = ?
                WHERE JobNumber = 'RB262';
                """

                # Execute the third chunk
                accessCursor.execute(updateQuery_Heading_3, *row[48:72])
                accessConn.commit()

                # Fourth chunk of fields
                updateQuery_Heading_4 = """
                UPDATE Heading
                SET InvoiceAddress1 = ?, InvoiceAddress2 = ?, InvoiceAddress3 = ?, InvoiceAddress4 = ?, InvoiceCounty = ?,
                    InvoiceNumber = ?, InvoicePostCode = ?, InvoicePrinted = ?, InvoiceSettled = ?, InvoiceSettledInformation = ?,
                    JobKeyID = ?, JobNumber = ?, JobType = ?, MainColour = ?, MasterJobKeyID = ?, Mobile = ?, Name = ?,
                    NettDeliveryCharge = ?, NettDeliveryCharge_Converted = ?, Notes = ?, PanelSupplierID = ?, PaymentMethod = ?,
                    PaymentMethodID = ?, PhoneNo = ?, Posted = ?, PreOrderPrinted = ?, PreQuotePrinted = ?, PriceGrids = ?,
                    Projectname = ?, QuantityFabricationUnits = ?, QuantityFabricationUnits_Fabrication = ?, QuantityFrames = ?
                WHERE JobNumber = 'RB262';
                """

                # Execute the fourth chunk
                accessCursor.execute(updateQuery_Heading_4, *row[72:96])
                accessConn.commit()

                # Fifth chunk of fields
                updateQuery_Heading_5 = """
                UPDATE Heading
                SET QuantityGlass = ?, QuantityGlassAnnealed = ?, QuantityGlassLaminated = ?, QuantityGlassObscure = ?,
                    QuantityGlassToughened = ?, QuantityPanels = ?, QuantityPanelsFlat = ?, QuantityPanelsMoulded = ?,
                    QuantityRoofs = ?, QuantityRoofPacks = ?, QuantityUnglazed = ?, QuoteNumber = ?, QuotePrinted = ?,
                    ReceivedOrder = ?, Reference = ?, RouteKeyID = ?, Salesman = ?, SalesmanCommission = ?, SalesmanID = ?,
                    Salutation = ?, SatelliteCustomerID = ?, SatelliteJobKeyID = ?, SatelliteJobNumber = ?, SatelliteName = ?,
                    SatelliteUploadedDate = ?, SellingPriceExTax = ?, SellingPriceExTax_Converted = ?, SellingPriceIncTax = ?,
                    SellingPriceIncTax_Converted = ?, TaxAmount = ?, TaxAmount_Converted = ?, TaxCode = ?, TaxRate = ?,
                    Toughened = ?, VersionCreated = ?, VersionAmended = ?
                WHERE JobNumber = 'RB262';
                """

                # Execute the fifth chunk
                accessCursor.execute(updateQuery_Heading_5, *row[96:])
                accessConn.commit()

                print("Access database updated successfully.")
            else:
                print("No data to update in Access database.")

except pyodbc.Error as ex:
    print(f"Access Database connection or update failed: {ex}")

finally:
    accessCursor.close()
    accessConn.close()
    print("Access Connection Closed.")
