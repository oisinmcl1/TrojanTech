import pyodbc


"""
CONNECT TO SQL SERVER
FETCH DBO.HEADING, DBO.HEADING2 AND DBO.HEADINGEVONET
"""

dsn = "Orders"
database = "FFOrders"

# Try connect to database and create a cursor
try:
    sqlServer = pyodbc.connect(f"DSN={dsn};DATABASE={database};")

    sqlCursor = sqlServer.cursor()

    print(f"Connection to SQL Server: {dsn} and Database: {database} was successful!")

    # Fetch rows from dbo_Heading where SatelliteJobNumber = 'RB262'
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

    # Verify fetched data
    if row:
        print("Fetched data:")
        print(row)
    else:
        print("No data found for SatelliteJobNumber = 'RB262'.")

# If error, print message and ensure rows is none
except pyodbc.Error as ex:
    print(f"SQL Server connection or query failed: {ex}")
    row = None


# Close connection
finally:
    sqlCursor.close()
    sqlServer.close()
    print("SQL Connection Closed.")





"""
CONNECT TO ACCESS DATABASE
UPDATE HEADING, HEADING2 AND HEADINGEVONET
"""

mdb = r"E:\BM\FuturaFrames\Orders\2024\Jun\Test\RB262.mdb"

try:
    # Connection string to Access database
    connStr = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={mdb};"
    )

    # Try to connect to Access database and create cursor
    accessConn = pyodbc.connect(connStr)

    accessCursor = accessConn.cursor()

    print(f"Connection to Access Database: {mdb} was successful!")

    if row:
        # Update Heading table in Access where Reference = 'RB2362'
        updateQuery_Heading = """
        UPDATE Heading
        SET AccountNo = ?, AcknowledgementStatus = ?, AmountPaid = ?,
            AmountPaid_Converted = ?, Architect = ?, Archive = ?, ArchivedCopyKeyID = ?, ArchiveZip = ?,
            BatchIndexOrder = ?, BatchKeyID = ?, BMGlassJob = ?, BrickCavity = ?, BuyUnglazedFrameGlass = ?,
            BuyUnglazedFromSupplier = ?, BuyPanelsFrom3rdParty = ?, ChangeDate = ?, ChangeTime = ?, CheckedOutTo = ?,
            CheckOrderCreated = ?, Completed = ?, Contact = ?, Cost = ?, CostLibrary = ?, CounterUniqueKeyID = ?,
            CreatedFrom = ?, CustomerID = ?, CustomerType = ?, CustomerUseMarkupPerColour = ?,
            CustomerUsePrimeAperture = ?, CustomerMidrailHeight = ?, CustomerTransomDrop = ?, DateAmended = ?,
            DateCompleted = ?, DateConfirmed = ?, DateCreated = ?, DateDelivery = ?, DateFitted = ?, DateGlassDelivery = ?,
            DateInProduction = ?, DateOnwardDelivery = ?, DateInvoice = ?, DateLoaded = ?, DateOrder = ?, DatePaid = ?,
            DatePaidSeveral = ?, DatePanelDelivery = ?, DateRoofDelivery = ?, DateSawFileCreated = ?, DateSurvey = ?,
            DateToughGlassDelivery = ?, DatePreOrderConverted = ?, DatePreQuoteConverted = ?, Despatched = ?,
            Dessian_CurrencyConversion = ?, Email = ?, EmailLong = ?, FaxNo = ?, FensaNumber = ?, FileName = ?, Fitting = ?,
            FittingAdvancedID = ?, FittingTeam = ?, FittingTeamID = ?, GlassAccountNo = ?, GlassBatched = ?,
            GlassSupplierID = ?, Glazed = ?, Hidden = ?, HousetypeCategoryID = ?, Initials = ?, InputBy = ?, InputByID = ?,
            InvoiceAddress1 = ?, InvoiceAddress2 = ?, InvoiceAddress3 = ?, InvoiceAddress4 = ?, InvoiceCounty = ?,
            InvoiceNumber = ?, InvoicePostCode = ?, InvoicePrinted = ?, InvoiceSettled = ?, InvoiceSettledInformation = ?,
            JobKeyID = ?, JobNumber = ?, JobType = ?, MainColour = ?, MasterJobKeyID = ?, Mobile = ?, Name = ?,
            NettDeliveryCharge = ?, NettDeliveryCharge_Converted = ?, Notes = ?, PanelSupplierID = ?, PaymentMethod = ?,
            PaymentMethodID = ?, PhoneNo = ?, Posted = ?, PreOrderPrinted = ?, PreQuotePrinted = ?, PriceGrids = ?,
            Projectname = ?, QuantityFabricationUnits = ?, QuantityFabricationUnits_Fabrication = ?, QuantityFrames = ?,
            QuantityGlass = ?, QuantityGlassAnnealed = ?, QuantityGlassLaminated = ?, QuantityGlassObscure = ?,
            QuantityGlassToughened = ?, QuantityPanels = ?, QuantityPanelsFlat = ?, QuantityPanelsMoulded = ?,
            QuantityRoofs = ?, QuantityRoofPacks = ?, QuantityUnglazed = ?, QuoteNumber = ?, QuotePrinted = ?,
            ReceivedOrder = ?, Reference = ?, RouteKeyID = ?, Salesman = ?, SalesmanCommission = ?, SalesmanID = ?,
            Salutation = ?, SatelliteCustomerID = ?, SatelliteJobKeyID = ?, SatelliteJobNumber = ?, SatelliteName = ?,
            SatelliteUploadedDate = ?, SellingPriceExTax = ?, SellingPriceExTax_Converted = ?, SellingPriceIncTax = ?,
            SellingPriceIncTax_Converted = ?, TaxAmount = ?, TaxAmount_Converted = ?, TaxCode = ?, TaxRate = ?,
            Toughened = ?, VersionCreated = ?, VersionAmended = ?
        WHERE JobNumber = 'RB262';
        """

        # Execute the update query with the data fetched from SQL Server
        accessCursor.execute(updateQuery_Heading, *row)

        # Commit the transaction
        accessConn.commit()

        print("Access database updated successfully.")

# If error, print error message
except pyodbc.Error as ex:
    print(f"Access Database connection failed: {ex}")

# Close connection
finally:
    try:
        accessCursor.close()
        accessConn.close()
        print("Access Connection Closed.")
    except:
        print("No connection to close or error closing connection")
