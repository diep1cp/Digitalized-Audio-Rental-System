# Digitalized Audio Rental System
This project aims to create a user-friendly, SQL-powered database solution to manage customer information, product inventory, and sales transactions.

#Log In – Design View

<img width="265" alt="image" src="https://github.com/user-attachments/assets/5de10838-384c-4574-b93c-500c13c69a4d">

#Log In – Execution Views

Option 1: Login as Manager

<img width="269" alt="image" src="https://github.com/user-attachments/assets/30107e1c-d60a-487e-888d-d56297cb94de">

Option 2: Login as Cashier

<img width="267" alt="image" src="https://github.com/user-attachments/assets/f406b5be-2eb6-42a8-b882-096b7ba32d52">

#Log In View Code

    Imports System.ComponentModel
    Public Class LogInForm
    Private DB As New DBAccess
    
    Private Sub LogInButton_Click(sender As Object, e As EventArgs) Handles LogInButton.Click
        EmployeeType = String.Empty
        If ValidateLoginCredentials() = False Then
            Exit Sub
        End If
        DB.AddParam("@username", UserIDTextBox.Text)
        DB.AddParam("@userpassword", UserPasswordTextBox.Text)
        DB.ExecuteQuery("SELECT * FROM Employee WHERE employeeid = ? AND passphrase = ? AND ActiveYN = 'Y'")
 
        If DB.DBException <> String.Empty Then
 
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
        If DB.RecordCount > 0 Then
 
            If DB.DBDataTable(0)!emptype = "A" Then
                EmployeeType = "A"
                MessageBox.Show("Successful Login as Manager")
 
            ElseIf DB.DBDataTable(0)!emptype = "B" Then
                EmployeeType = "B"
                MessageBox.Show("Successful Login as Cashier")
 
            End If
            Homepage.ShowDialog()
            Me.Close()
        End If
 
    End Sub
    Private Sub LogInForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
 
        UserIDTextBox.Clear()
        UserPasswordTextBox.Clear()
 
    End Sub
 
    Private Function ValidateLoginCredentials() As Boolean
 
        If String.IsNullOrWhiteSpace(UserIDTextBox.Text) Then
            MessageBox.Show("User ID cannot be empty.")
            UserIDTextBox.SelectAll()
            UserIDTextBox.Focus()
            Return False
        End If
        If String.IsNullOrWhiteSpace(UserPasswordTextBox.Text) Then
            MessageBox.Show("User Password cannot be empty.")
            UserPasswordTextBox.SelectAll()
            UserPasswordTextBox.Focus()
            Return False
        End If
 
        Return True
    End Function
    Private Function AuthorizeLogin() As Boolean
 
        Dim ExactPasswordString As String
 
        DB.AddParam("@username", UserIDTextBox.Text)
        DB.AddParam("@userpassword", UserPasswordTextBox.Text)
        DB.ExecuteQuery("SELECT * FROM Employee WHERE employeeid = ? AND passphrase = ? AND ActiveYN = Y")
 
        If DB.DBException <> String.Empty Then
            Return False
        End If
 
 
        If DB.RecordCount > 0 Then
            Return True
        End If
 
    End Function
 
    Private Sub LogInForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
 
    End Sub

    
    End Class

Home Page View

Home Page - Design View

<img width="432" alt="image" src="https://github.com/user-attachments/assets/9d65ddcd-3185-4b25-88dd-3e85ea2c4900">

Home Page - Execution View

Click customer form to begin

<img width="455" alt="image" src="https://github.com/user-attachments/assets/bcec970e-1089-4f86-855f-f2ebd9b8a544">

Home Page View Code

    Public Class Homepage
    Private DB As New DBAccess
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles InventoryButton.Click
        ManagerForm.Show()
    End Sub 
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles CustomersButton.Click
        Customers.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles OrdersButton.Click
        RegularOrder.Show()
 
    End Sub
 
    Private Sub ServiceButton_Click(sender As Object, e As EventArgs) Handles ServiceButton.Click
        NewServiceOrder.Show()
    End Sub 
    
    Private Sub Homepage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If EmployeeType = "B" Then
            InventoryButton.Visible = False
        End If
    End Sub

    End Class

Customer Form Views

Customer Form - Design Views

<img width="468" alt="image" src="https://github.com/user-attachments/assets/b96f8f51-16a8-4ebb-a5ca-059578d74b4f">

Customer Form – Execution Views

Step 1: The form appears when you select the “Customer” Form.

<img width="468" alt="image" src="https://github.com/user-attachments/assets/dce7974e-55f6-4cfa-a1f6-bbdba4021916">

Step 2: Input “J” in the First Name Text Box

<img width="465" alt="image" src="https://github.com/user-attachments/assets/2709e52f-739a-48e6-a322-ec24f75d4964">

Step 3: Input “D” in the Last Name Text Box

<img width="464" alt="image" src="https://github.com/user-attachments/assets/0f3b5f49-f190-4233-8e2c-099637e78073">

Step 4: Input “C” in the First Name Text Box and “W” in the Last Name Text Box

<img width="461" alt="image" src="https://github.com/user-attachments/assets/cfbdff96-0e42-47ed-9089-eebbbe39a832">

Step 5: If the customer is not in the database, the user can create a new customer by adding information such as First Name, Last Name, Phone, Email.

<img width="297" alt="image" src="https://github.com/user-attachments/assets/82278279-7222-4bd4-8ec1-f4a742a2ba96">

Step 6 : Once everything is completed, select “Create New Customer” to receive the pop-up message confirming the new customer is added successfully.

<img width="449" alt="image" src="https://github.com/user-attachments/assets/c852702d-8a47-47e0-961f-9dc3c3bbde79">

Customer Form Code 

    Imports System.Text.RegularExpressions
    Public Class Customers
    Private DB As New DBAccess
    Private Sub Customers_Load(sender As Object, e As EventArgs) Handles Me.Load
        SearchCustomer("%", "%")
    End Sub


    Private Sub SearchCustomer(FirstName As String, LastName As String)
        DB.AddParam("@firstname", FirstName & "%")
        DB.AddParam("@lastname", LastName & "%")
        DB.ExecuteQuery("SELECT * FROM customer WHERE firstname LIKE ? AND lastname LIKE ?")
 
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
        DBDataGridView.DataSource = DB.DBDataTable
 
    End Sub
 
    Private Sub FirstNameText_KeyUp(sender As Object, e As KeyEventArgs) Handles FirstNameText.KeyUp
        SearchCustomer(FirstNameText.Text, LastNameText.Text)
    End Sub
 
    Private Sub LastNameText_KeyUp(sender As Object, e As KeyEventArgs) Handles LastNameText.KeyUp
        SearchCustomer(FirstNameText.Text, LastNameText.Text)
    End Sub
 
    Private Sub AddNewCustomer()
        DB.AddParam("@firstname", FirstNameTextBox.Text)
        DB.AddParam("@lastname", LastNameTextBox.Text)
        DB.AddParam("phonenum", PhoneTextBox.Text)
        DB.AddParam("@create_date", Today.Date.ToString("yyyy-MM-dd"))
        DB.AddParam("@last_update", Today.Date.ToString("yyyy-MM-dd"))
        DB.AddParam("@activeYN", "Y")
        If String.IsNullOrEmpty(EmailTextBox.Text) Then
            DB.AddParam("@email", DBNull.Value)
        Else
            DB.AddParam("@email", EmailTextBox.Text) '3rd param, for third ? 
        End If
 
        DB.ExecuteQuery("INSERT INTO customer(FirstName, LastName, PhoneNum, Create_Date, Last_Update, ActiveYN, Email) VALUES(?, ?, ?, ?, ?, ?, ?)")
 
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
        MessageBox.Show("A new customer " & FirstNameTextBox.Text & " " & LastNameTextBox.Text & " is added successfully.")
 
    End Sub
 
 
    Private Function ValidNewCustomer() As Boolean
        'validate the store_id field
        'step 1 - make sure the first name field is not empty
        If String.IsNullOrWhiteSpace(FirstNameTextBox.Text) = True Then
            MessageBox.Show("The store ID can't be empty.")
            FirstNameTextBox.SelectAll()
            FirstNameTextBox.Focus()
            Return False
        End If
 
        'step 2 - checks if the first name meets the right pattern
        Dim rxFirstName As New Regex("^[A-Z]+[a-z]+$")
        If rxFirstName.IsMatch(FirstNameTextBox.Text) = False Then
            MessageBox.Show("The first name must be alphabetical and start with an uppercase letter.")
            FirstNameTextBox.SelectAll()
            FirstNameTextBox.Focus()
            Return False
        End If
 
        '------------------------------------------------------------------------
 
        'validate last name field
        'step 1 - check if the last name is empty
        If String.IsNullOrWhiteSpace(LastNameTextBox.Text) Then
            MessageBox.Show("The last name can't be empty.")
            LastNameTextBox.SelectAll()
            LastNameTextBox.Focus()
 
            Return False
        End If
 
        'step 2 - checks if the last name meets the right pattern
        Dim rxLastName As New Regex("^[A-Z]+[a-z]+$")
        If rxLastName.IsMatch(LastNameTextBox.Text) = False Then
            MessageBox.Show("The last name must be alphabetical and start with an uppercase letter.")
            LastNameTextBox.SelectAll()
            LastNameTextBox.Focus()
 
            Return False
        End If
 
        '------------------------------------------------------------------------
 
        'validate the phone number field
        'step 1 - make sure the phone number field is not empty
        If String.IsNullOrWhiteSpace(PhoneTextBox.Text) = True Then
            MessageBox.Show("The phone number can't be empty.")
            PhoneTextBox.SelectAll()
            PhoneTextBox.Focus()
 
            Return False
        End If
 
        '------------------------------------------------------------------------
 
        'validate the email field
        'check if empty
        If String.IsNullOrEmpty(EmailTextBox.Text) = True Then
            'do nothing
        Else
            'validate email entry
            Dim rxEmail As New Regex("[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
            If rxEmail.IsMatch(EmailTextBox.Text) = False Then
                MessageBox.Show("Email must follow the right format such as abc@xyz.com.")
                EmailTextBox.SelectAll()
                EmailTextBox.Focus()
                Return False
            End If
        End If
 
        Return True
    End Function
 
 
    Private Function DuplicateCustomer() As Boolean
        DB.AddParam("@firstname", FirstNameTextBox.Text)
        DB.AddParam("@lastname", LastNameTextBox.Text)
 
        DB.ExecuteQuery("SELECT * FROM customer WHERE first_name = ? AND last_name = ?")
 
        If DB.RecordCount > 0 Then
            'found an existing record matching the three conditions
            MessageBox.Show("The same customer already exists")
            Return True
        End If
 
        Return False
    End Function
 
    Private Sub CreateCustomer_Click(sender As Object, e As EventArgs) Handles CreateCustomerButton.Click
        If ValidNewCustomer() = True Then
            If DuplicateCustomer() = False Then
                AddNewCustomer()
                Me.Close()
            End If
        End If
    End Sub
    
    End Class

Manager/Inventory – Design Views

<img width="468" alt="image" src="https://github.com/user-attachments/assets/3061fc9b-298d-4826-b894-38427957c4b5">

Manager/Inventory – Execution Views

<img width="468" alt="image" src="https://github.com/user-attachments/assets/3b825dc3-9d19-4120-b4e0-42028f144d4f">

Manager/Inventory Code

    Public Class ManagerForm
        Dim itemname 'declaration
        
    Private DB As New DBAccess
    Private Sub ManagerForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SearchItem("%")
        DB.ExecuteQuery("SELECT * FROM item WHERE ActiveYN = 'Y'") 'selects the active inventory items from the SQL item table
 
        If DB.DBException <> String.Empty Then 'error check
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
        'data source
        ItemListBox.DataSource = DB.DBDataTable
        ItemListBox.DisplayMember = "itemname"
 
    End Sub
 
 
    Private Sub ItemListBox_DoubleClick(sender As Object, e As EventArgs) Handles ItemListBox.DoubleClick 'when item list box is double-clicked
        Dim CompleteRemove As DialogResult
        CompleteRemove = MessageBox.Show("Are you sure you want to remove " & ItemListBox.GetItemText(ItemListBox.SelectedItem) & " from inventory?")
        If CompleteRemove = vbOK Then
            DB.AddParam("@itemname", ItemListBox.GetItemText(ItemListBox.SelectedItem))
            DB.ExecuteQuery("Update item Set ActiveYN = 'N' WHERE ItemName = ?") 'sets selected item as inactive 
            If DB.DBException <> String.Empty Then
                MessageBox.Show(DB.DBException)
                Exit Sub
            End If
            'display message to user that item has been removed
            MessageBox.Show("The item " & ItemListBox.SelectedItem!itemname & " has been successfully removed from inventory.")
        End If
    End Sub
 
 
    Private Sub AddNewItem()
        DB.AddParam("@itemname", ItemNameTextBox.Text) 'first paramater, for 1st ?
        DB.AddParam("itemdescription", ItemDescriptionTextBox.Text) 'secnond parameter, for 2nd ?
        DB.AddParam("@price", PriceTextBox.Text) 'third parameter, for 3rd ?
        DB.AddParam("@qtyavailable", QuantityTextBox.Text) 'fourth parameter, for 4th ?
        DB.AddParam("@create_date", Today.Date.ToString("yyyy-MM-dd")) 'fifth parameter, for 5th ?
        DB.AddParam("@last_update", Today.Date.ToString("yyyy-MM-dd")) 'sixth parameter, for 6th ?
 
        'inserts new record into item table containing user entry
        DB.ExecuteQuery("INSERT INTO item(ItemName, ItemDescription, Price, QtyAvailable, Create_Date, Last_Update, ActiveYN) VALUES(?, ?, ?, ?, ?, ?, 'Y')")
 
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
        'Display message to user 
        MessageBox.Show("A new item " & ItemNameTextBox.Text & " is added successfully.")
 
    End Sub
 
 
    Private Sub SearchItem(ItemName As String)
        DB.AddParam("@itemname", ItemName & "%")
        DB.ExecuteQuery("SELECT * FROM item WHERE itemname LIKE ?")
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
    End Sub
 
    'Validations
    Private Function ValidNewItem() As Boolean
        'validate the item name field
        'step 1 - check if the item name is empty
        If String.IsNullOrWhiteSpace(ItemNameTextBox.Text) Then
            MessageBox.Show("The item name can't be empty.")
            ItemNameTextBox.SelectAll()
            ItemNameTextBox.Focus()
            Return False
        End If
 
        '------------------------------------------------------------------------
        'validate the item description field
        'step 1 - check if the item description is empty
        If String.IsNullOrWhiteSpace(ItemDescriptionTextBox.Text) Then
            MessageBox.Show("The item description can't be empty.")
            ItemDescriptionTextBox.SelectAll()
            ItemDescriptionTextBox.Focus()
            Return False
        End If
 
        '------------------------------------------------------------------------
        'validate the price field
        'step 1 - check if the price is empty
        If String.IsNullOrWhiteSpace(PriceTextBox.Text) Then
            MessageBox.Show("The price can't be empty.")
            PriceTextBox.SelectAll()
            PriceTextBox.Focus()
            Return False
        End If
 
        'step 2 - make sure the price is a decimal.
        Dim PriceDec As Decimal
        If Decimal.TryParse(PriceTextBox.Text, PriceDec) = False Then
            MessageBox.Show("The price must be a decimal.")
            PriceTextBox.SelectAll()
            PriceTextBox.Focus()
            Return False
        End If
 
        '------------------------------------------------------------------------
        'validate the quantity field
        'step 1 - check if the quantity is empty
        If String.IsNullOrWhiteSpace(QuantityTextBox.Text) Then
            MessageBox.Show("The quantity can't be empty.")
            PriceTextBox.SelectAll()
            PriceTextBox.Focus()
            Return False
        End If
 
        'step 2 - make sure the quantity is an integer.
        Dim QuantityInteger As Integer
        If Integer.TryParse(QuantityTextBox.Text, QuantityInteger) = False Then
            MessageBox.Show("The quantity must be an integer.")
            QuantityTextBox.SelectAll()
            QuantityTextBox.Focus()
            Return False
        End If
 
        '3 - make sure the quantity has the right value from the parent table Store
        If QuantityInteger < 1 Or QuantityInteger > 50 Then
            MessageBox.Show("The quantity must be between 1 and 50.")
            QuantityTextBox.SelectAll()
            QuantityTextBox.Focus()
            Return False
        End If
 
        Return True
    End Function
 
    'Checks if the item already exists in the database
    Private Function DuplicateItem() As Boolean
        DB.AddParam("@itemname", ItemNameTextBox.Text)
 
        DB.ExecuteQuery("SELECT * FROM item WHERE itemname = ?")
 
        If DB.RecordCount > 0 Then
            'found an existing record matching the three conditions
            MessageBox.Show("The same item already exists")
            Return True
        End If
 
        Return False
    End Function
 
    Private Sub AddNewItemButton_Click(sender As Object, e As EventArgs) Handles AddNewItemButton.Click
        If ValidNewItem() = True Then 'checks if item validation is confirmed
            If DuplicateItem() = False Then 'checks if no duplicate items exist
                AddNewItem() 'executes query to add item
                Me.Close()
            End If
        End If
    End Sub


New Service Order View

New Service Order – Design Views

<img width="468" alt="image" src="https://github.com/user-attachments/assets/a0b6bb8e-9f09-47d9-a7a0-2e81eb675a63">

New Service Order – Execution Views

<img width="468" alt="image" src="https://github.com/user-attachments/assets/ef73cca7-e943-425e-97a4-00a5a07bc519">


•	This module allows for new service orders to be added to the database.
•	The required criteria for a new service order are Items Needed to Repair, Select Customer, Select Employee, ID, Item Name, Unit Price, Quantity, Sub Total, Start Date, Finish Date, and Repair Description.  

•	Select an item from the list that needs to be repaired and select the customer and an employee who will be performing the repair. The system will automatically populate the ID, Item Name, calculate the Unit Price for each Quantity, and provide the total amount the customer needs to pay to get this service. 

•	Select the Start Date and the Finish Date

<img width="458" alt="image" src="https://github.com/user-attachments/assets/e418e687-1704-4f09-808d-2b1a789df2a0">

•	    Type in Repair Description

<img width="443" alt="image" src="https://github.com/user-attachments/assets/a2afeaec-c6ba-436c-a278-5009eb336365">

*Select “Send Repair” and receive a pop-up message that shows the order summary

<img width="435" alt="image" src="https://github.com/user-attachments/assets/8d6fd521-26ac-469c-80af-eae1731d245c">

New Service Order Code 

    Public Class ManagerForm
        Dim itemname 'declaration
      
    Private DB As New DBAccess
    Private Sub ManagerForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SearchItem("%")
        DB.ExecuteQuery("SELECT * FROM item WHERE ActiveYN = 'Y'") 'selects the active inventory items from the SQL            item table
 
        If DB.DBException <> String.Empty Then 'error check
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
        'data source
        ItemListBox.DataSource = DB.DBDataTable
        ItemListBox.DisplayMember = "itemname"
 
    End Sub
 
 
    Private Sub ItemListBox_DoubleClick(sender As Object, e As EventArgs) Handles ItemListBox.DoubleClick 'when item list box is double-clicked
        Dim CompleteRemove As DialogResult
        CompleteRemove = MessageBox.Show("Are you sure you want to remove " & ItemListBox.GetItemText(ItemListBox.SelectedItem) & " from inventory?")
        If CompleteRemove = vbOK Then
            DB.AddParam("@itemname", ItemListBox.GetItemText(ItemListBox.SelectedItem))
            DB.ExecuteQuery("Update item Set ActiveYN = 'N' WHERE ItemName = ?") 'sets selected item as inactive 
            If DB.DBException <> String.Empty Then
                MessageBox.Show(DB.DBException)
                Exit Sub
            End If
            'display message to user that item has been removed
            MessageBox.Show("The item " & ItemListBox.SelectedItem!itemname & " has been successfully removed from inventory.")
        End If
    End Sub
 
 
    Private Sub AddNewItem()
        DB.AddParam("@itemname", ItemNameTextBox.Text) 'first paramater, for 1st ?
        DB.AddParam("itemdescription", ItemDescriptionTextBox.Text) 'secnond parameter, for 2nd ?
        DB.AddParam("@price", PriceTextBox.Text) 'third parameter, for 3rd ?
        DB.AddParam("@qtyavailable", QuantityTextBox.Text) 'fourth parameter, for 4th ?
        DB.AddParam("@create_date", Today.Date.ToString("yyyy-MM-dd")) 'fifth parameter, for 5th ?
        DB.AddParam("@last_update", Today.Date.ToString("yyyy-MM-dd")) 'sixth parameter, for 6th ?
 
        'inserts new record into item table containing user entry
        DB.ExecuteQuery("INSERT INTO item(ItemName, ItemDescription, Price, QtyAvailable, Create_Date, Last_Update, ActiveYN) VALUES(?, ?, ?, ?, ?, ?, 'Y')")
 
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
        'Display message to user 
        MessageBox.Show("A new item " & ItemNameTextBox.Text & " is added successfully.")
 
    End Sub
 
 
    Private Sub SearchItem(ItemName As String)
        DB.AddParam("@itemname", ItemName & "%")
        DB.ExecuteQuery("SELECT * FROM item WHERE itemname LIKE ?")
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
    End Sub
 
    'Validations
    Private Function ValidNewItem() As Boolean
        'validate the item name field
        'step 1 - check if the item name is empty
        If String.IsNullOrWhiteSpace(ItemNameTextBox.Text) Then
            MessageBox.Show("The item name can't be empty.")
            ItemNameTextBox.SelectAll()
            ItemNameTextBox.Focus()
            Return False
        End If
 
        '------------------------------------------------------------------------
        'validate the item description field
        'step 1 - check if the item description is empty
        If String.IsNullOrWhiteSpace(ItemDescriptionTextBox.Text) Then
            MessageBox.Show("The item description can't be empty.")
            ItemDescriptionTextBox.SelectAll()
            ItemDescriptionTextBox.Focus()
            Return False
        End If
 
        '------------------------------------------------------------------------
        'validate the price field
        'step 1 - check if the price is empty
        If String.IsNullOrWhiteSpace(PriceTextBox.Text) Then
            MessageBox.Show("The price can't be empty.")
            PriceTextBox.SelectAll()
            PriceTextBox.Focus()
            Return False
        End If
 
        'step 2 - make sure the price is a decimal.
        Dim PriceDec As Decimal
        If Decimal.TryParse(PriceTextBox.Text, PriceDec) = False Then
            MessageBox.Show("The price must be a decimal.")
            PriceTextBox.SelectAll()
            PriceTextBox.Focus()
            Return False
        End If
 
        '------------------------------------------------------------------------
        'validate the quantity field
        'step 1 - check if the quantity is empty
        If String.IsNullOrWhiteSpace(QuantityTextBox.Text) Then
            MessageBox.Show("The quantity can't be empty.")
            PriceTextBox.SelectAll()
            PriceTextBox.Focus()
            Return False
        End If
 
        'step 2 - make sure the quantity is an integer.
        Dim QuantityInteger As Integer
        If Integer.TryParse(QuantityTextBox.Text, QuantityInteger) = False Then
            MessageBox.Show("The quantity must be an integer.")
            QuantityTextBox.SelectAll()
            QuantityTextBox.Focus()
            Return False
        End If
 
        '3 - make sure the quantity has the right value from the parent table Store
        If QuantityInteger < 1 Or QuantityInteger > 50 Then
            MessageBox.Show("The quantity must be between 1 and 50.")
            QuantityTextBox.SelectAll()
            QuantityTextBox.Focus()
            Return False
        End If
 
        Return True
    End Function
 
    'Checks if the item already exists in the database
    Private Function DuplicateItem() As Boolean
        DB.AddParam("@itemname", ItemNameTextBox.Text)
 
        DB.ExecuteQuery("SELECT * FROM item WHERE itemname = ?")
 
        If DB.RecordCount > 0 Then
            'found an existing record matching the three conditions
            MessageBox.Show("The same item already exists")
            Return True
        End If
 
        Return False
    End Function
 
    Private Sub AddNewItemButton_Click(sender As Object, e As EventArgs) Handles AddNewItemButton.Click
        If ValidNewItem() = True Then 'checks if item validation is confirmed
            If DuplicateItem() = False Then 'checks if no duplicate items exist
                AddNewItem() 'executes query to add item
                Me.Close()
            End If
        End If
    End Sub
End Class

New Service Order Code – DBAccess.vb

    Public Class NewServiceOrder
        Private DB As New DBAccess
        Private Sub NewServiceOrder_Load(sender As Object, e As EventArgs) Handles Me.Load
              DB.ExecuteQuery("SELECT * FROM item WHERE itemid>20")
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
        RepairItemListBox.DataSource = DB.DBDataTable
        RepairItemListBox.DisplayMember = "itemname"
 
        DB.ExecuteQuery("SELECT * FROM customer")
 
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
        CustomerListBox.DataSource = DB.DBDataTable
        CustomerListBox.DisplayMember = "firstname"
        CustomerListBox.ValueMember = "customerid"
 
        DB.ExecuteQuery("SELECT * FROM employee")
 
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
        EmployeeListBox.DataSource = DB.DBDataTable
        EmployeeListBox.DisplayMember = "firstname"
        EmployeeListBox.ValueMember = "employeeid"
    End Sub
 
    Private Sub RepairItemListBox_DoubleClick(sender As Object, e As EventArgs) Handles RepairItemListBox.DoubleClick
 
        Dim theIndex As Integer = ItemNameListBox.FindString(RepairItemListBox.SelectedItem!itemname)
        If theIndex = -1 Then
            ItemNameListBox.Items.Add(RepairItemListBox.SelectedItem!itemname)
            RepairItemIDListBox.Items.Add(RepairItemListBox.SelectedItem!itemID)
            UnitPriceListBox.Items.Add(RepairItemListBox.SelectedItem!price)
            QuantityListBox.Items.Add("1")
            SubtotalListBox.Items.Add(RepairItemListBox.SelectedItem!price * 1)
 
        Else
            QuantityListBox.Items(theIndex) += 1
            SubtotalListBox.Items(theIndex) = QuantityListBox.Items(theIndex) * UnitPriceListBox.Items(theIndex)
        End If
 
        Dim ordertotal As Decimal
        For Each item In SubtotalListBox.Items
            ordertotal = ordertotal + item
        Next
        RepairOrderTotalTextBox.Text = ordertotal.ToString("C")
        TaxTextBox.Text = (ordertotal * 0.06).ToString("C")
        RepairFeeTextBox.Text = "$50.00"
        TotalTextBox.Text = (ordertotal + TaxTextBox.Text + RepairFeeTextBox.Text).ToString("C")
 
    End Sub
 
    Private Sub ItemNameListBox_DoubleClick(sender As Object, e As EventArgs) Handles ItemNameListBox.DoubleClick
        Dim theIndex As Integer = ItemNameListBox.SelectedIndex
 
        If QuantityListBox.Items(theIndex) = 1 Then
            RepairItemIDListBox.Items.RemoveAt(theIndex)
            ItemNameListBox.Items.RemoveAt(theIndex)
            UnitPriceListBox.Items.RemoveAt(theIndex)
            QuantityListBox.Items.RemoveAt(theIndex)
            SubtotalListBox.Items.RemoveAt(theIndex)
        Else
            QuantityListBox.Items(theIndex) -= 1
            SubtotalListBox.Items(theIndex) -= UnitPriceListBox.Items(theIndex)
        End If
 
        Dim ordertotal As Decimal
        For Each item In SubtotalListBox.Items
            ordertotal = ordertotal + item
        Next
        RepairOrderTotalTextBox.Text = ordertotal.ToString("C")
        TaxTextBox.Text = (ordertotal * 0.06).ToString("C")
        TotalTextBox.Text = (ordertotal + TaxTextBox.Text).ToString("C")
    End Sub
 
    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        Me.Close()
    End Sub
 
    Public Sub SendRepairButton_Click(sender As Object, e As EventArgs) Handles SendRepairButton.Click
        Dim OrderSummary As String
 
        For i = 0 To ItemNameListBox.Items.Count - 1
            OrderSummary += QuantityListBox.Items(i) & " " & ItemNameListBox.Items(i)
            OrderSummary += vbCrLf
        Next
 
        OrderSummary += vbCrLf
        OrderSummary += "Repair Order Total: " & RepairOrderTotalTextBox.Text
        OrderSummary += vbCrLf
        OrderSummary += vbCrLf
        OrderSummary += "Tax: " & TaxTextBox.Text
        OrderSummary += vbCrLf
        OrderSummary += vbCrLf
        OrderSummary += "Repair Fee: " & RepairFeeTextBox.Text
        OrderSummary += vbCrLf
        OrderSummary += vbCrLf
        OrderSummary += "Total: " & TotalTextBox.Text
 
        Dim Complete As DialogResult
        Complete = MessageBox.Show("Does this complete the repair order?" & vbCrLf & vbCrLf & OrderSummary, "Order Summary", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
 
        Dim total As Double
        If Complete = vbOK Then
            total = TotalTextBox.Text.Replace("$", "")
            DB.AddParam("@create_date", StartTime.Value)
            DB.AddParam("@finish_date", FinishTime.Value)
            DB.AddParam("@repaircost", total)
            DB.AddParam("@repairdescription", RepairDescTextbox.Text)
            DB.AddParam("@customerid", CustomerListBox.SelectedValue)
            DB.AddParam("@employeeid", EmployeeListBox.SelectedValue)
            DB.ExecuteQuery("INSERT INTO repairorder(create_date, finish_date, repaircost, repairdescription, customerid, employeeid) VALUES (?,?,?,?,?,?)")
            If DB.DBException <> String.Empty Then
                MessageBox.Show(DB.DBException)
                Exit Sub
            End If
            For i = 0 To ItemNameListBox.Items.Count - 1
                ItemInserts(i)
            Next
            Me.Close()
        Else
            'do nothing
        End If
    End Sub
 
    Public Sub ItemInserts(value As Integer)
        Dim id As String
        DB.ExecuteQuery("SELECT max(repairorderid) from repairorder")
        id = DB.DBDataTable.Rows.Item(0).Item(0).ToString
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
        DB.AddParam("@repairorderid", id)
        DB.AddParam("@itemid", RepairItemIDListBox.Items(value))
        DB.AddParam("@quantity", QuantityListBox.Items(value))
        DB.AddParam("@price", SubtotalListBox.Items(value))
        DB.ExecuteQuery("INSERT INTO repairorderitem(repairorderid, itemid, quantity, price) VALUES (?,?,?,?)")
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
    End Sub
    
    End Class

Regular Order View

Regular Order View – Design Views

<img width="468" alt="image" src="https://github.com/user-attachments/assets/7857f49d-f3e5-47a5-880b-f661eed85822">

Regular Order View – Execution Views

Step 1: Double-click Customer name in Select Customer list box

<img width="360" alt="image" src="https://github.com/user-attachments/assets/e0ad524d-bf60-4af1-9228-555266a838dd">

Step 2: Double click Employee name in the Select Employee list box 

<img width="360" alt="image" src="https://github.com/user-attachments/assets/6cd18a82-b3b5-467d-a759-0c9a2fa025ac">

Step 3: Double-click item name in Add Items to Order listbox

<img width="360" alt="image" src="https://github.com/user-attachments/assets/e3ef4a96-fa54-46e2-8b1c-38c0a4bc328b">

Step 4: Double click item again to increase quantity 

<img width="360" alt="image" src="https://github.com/user-attachments/assets/4cd5670d-2a21-4a4e-86c3-b3129007a108">

Step 5: Double click item name in Items Added list box to reduce quantity 

<img width="360" alt="image" src="https://github.com/user-attachments/assets/828be7c8-c8ab-4325-96b4-14c435cd0bcf">

Regular Order View – Code

        Public Class NewServiceOrder
            Private DB As New DBAccess
            Private Sub NewServiceOrder_Load(sender As Object, e As EventArgs) Handles Me.Load
                DB.ExecuteQuery("SELECT * FROM item WHERE itemid>20")
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
        
        RepairItemListBox.DataSource = DB.DBDataTable
        RepairItemListBox.DisplayMember = "itemname"
 
        DB.ExecuteQuery("SELECT * FROM customer")
 
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
        CustomerListBox.DataSource = DB.DBDataTable
        CustomerListBox.DisplayMember = "firstname"
        CustomerListBox.ValueMember = "customerid"
 
        DB.ExecuteQuery("SELECT * FROM employee")
 
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
 
        EmployeeListBox.DataSource = DB.DBDataTable
        EmployeeListBox.DisplayMember = "firstname"
        EmployeeListBox.ValueMember = "employeeid"
    End Sub
 
    Private Sub RepairItemListBox_DoubleClick(sender As Object, e As EventArgs) Handles RepairItemListBox.DoubleClick
 
        Dim theIndex As Integer = ItemNameListBox.FindString(RepairItemListBox.SelectedItem!itemname)
        If theIndex = -1 Then
            ItemNameListBox.Items.Add(RepairItemListBox.SelectedItem!itemname)
            RepairItemIDListBox.Items.Add(RepairItemListBox.SelectedItem!itemID)
            UnitPriceListBox.Items.Add(RepairItemListBox.SelectedItem!price)
            QuantityListBox.Items.Add("1")
            SubtotalListBox.Items.Add(RepairItemListBox.SelectedItem!price * 1)
 
        Else
            QuantityListBox.Items(theIndex) += 1
            SubtotalListBox.Items(theIndex) = QuantityListBox.Items(theIndex) * UnitPriceListBox.Items(theIndex)
        End If
 
        Dim ordertotal As Decimal
        For Each item In SubtotalListBox.Items
            ordertotal = ordertotal + item
        Next
        RepairOrderTotalTextBox.Text = ordertotal.ToString("C")
        TaxTextBox.Text = (ordertotal * 0.06).ToString("C")
        RepairFeeTextBox.Text = "$50.00"
        TotalTextBox.Text = (ordertotal + TaxTextBox.Text + RepairFeeTextBox.Text).ToString("C")
 
    End Sub
 
    Private Sub ItemNameListBox_DoubleClick(sender As Object, e As EventArgs) Handles ItemNameListBox.DoubleClick
        Dim theIndex As Integer = ItemNameListBox.SelectedIndex
 
        If QuantityListBox.Items(theIndex) = 1 Then
            RepairItemIDListBox.Items.RemoveAt(theIndex)
            ItemNameListBox.Items.RemoveAt(theIndex)
            UnitPriceListBox.Items.RemoveAt(theIndex)
            QuantityListBox.Items.RemoveAt(theIndex)
            SubtotalListBox.Items.RemoveAt(theIndex)
        Else
            QuantityListBox.Items(theIndex) -= 1
            SubtotalListBox.Items(theIndex) -= UnitPriceListBox.Items(theIndex)
        End If
 
        Dim ordertotal As Decimal
        For Each item In SubtotalListBox.Items
            ordertotal = ordertotal + item
        Next
        RepairOrderTotalTextBox.Text = ordertotal.ToString("C")
        TaxTextBox.Text = (ordertotal * 0.06).ToString("C")
        TotalTextBox.Text = (ordertotal + TaxTextBox.Text).ToString("C")
    End Sub
 
    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        Me.Close()
    End Sub
 
    Public Sub SendRepairButton_Click(sender As Object, e As EventArgs) Handles SendRepairButton.Click
        Dim OrderSummary As String
 
        For i = 0 To ItemNameListBox.Items.Count - 1
            OrderSummary += QuantityListBox.Items(i) & " " & ItemNameListBox.Items(i)
            OrderSummary += vbCrLf
        Next
 
        OrderSummary += vbCrLf
        OrderSummary += "Repair Order Total: " & RepairOrderTotalTextBox.Text
        OrderSummary += vbCrLf
        OrderSummary += vbCrLf
        OrderSummary += "Tax: " & TaxTextBox.Text
        OrderSummary += vbCrLf
        OrderSummary += vbCrLf
        OrderSummary += "Repair Fee: " & RepairFeeTextBox.Text
        OrderSummary += vbCrLf
        OrderSummary += vbCrLf
        OrderSummary += "Total: " & TotalTextBox.Text
 
        Dim Complete As DialogResult
        Complete = MessageBox.Show("Does this complete the repair order?" & vbCrLf & vbCrLf & OrderSummary, "Order Summary", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
 
        Dim total As Double
        If Complete = vbOK Then
            total = TotalTextBox.Text.Replace("$", "")
            DB.AddParam("@create_date", StartTime.Value)
            DB.AddParam("@finish_date", FinishTime.Value)
            DB.AddParam("@repaircost", total)
            DB.AddParam("@repairdescription", RepairDescTextbox.Text)
            DB.AddParam("@customerid", CustomerListBox.SelectedValue)
            DB.AddParam("@employeeid", EmployeeListBox.SelectedValue)
            DB.ExecuteQuery("INSERT INTO repairorder(create_date, finish_date, repaircost, repairdescription, customerid, employeeid) VALUES (?,?,?,?,?,?)")
            If DB.DBException <> String.Empty Then
                MessageBox.Show(DB.DBException)
                Exit Sub
            End If
            For i = 0 To ItemNameListBox.Items.Count - 1
                ItemInserts(i)
            Next
            Me.Close()
        Else
            'do nothing
        End If
    End Sub
 
    Public Sub ItemInserts(value As Integer)
        Dim id As String
        DB.ExecuteQuery("SELECT max(repairorderid) from repairorder")
        id = DB.DBDataTable.Rows.Item(0).Item(0).ToString
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
        DB.AddParam("@repairorderid", id)
        DB.AddParam("@itemid", RepairItemIDListBox.Items(value))
        DB.AddParam("@quantity", QuantityListBox.Items(value))
        DB.AddParam("@price", SubtotalListBox.Items(value))
        DB.ExecuteQuery("INSERT INTO repairorderitem(repairorderid, itemid, quantity, price) VALUES (?,?,?,?)")
        If DB.DBException <> String.Empty Then
            MessageBox.Show(DB.DBException)
            Exit Sub
        End If
    End Sub
    
    End Class






