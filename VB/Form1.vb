Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditCustomizeMergeFields

    Public Partial Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
            Dim invoices As List(Of Invoice) = New List(Of Invoice)(10)
            invoices.Add(New Invoice(0, "Invoice1", 10.0D, Guid.NewGuid()))
            invoices.Add(New Invoice(1, "Invoice2", 15.0D, Guid.NewGuid()))
            invoices.Add(New Invoice(2, "Invoice3", 20.0D, Guid.NewGuid()))
            richEditControl1.Options.MailMerge.DataSource = invoices
        End Sub

        Private Sub richEditControl1_CustomizeMergeFields(ByVal sender As Object, ByVal e As DevExpress.XtraRichEdit.CustomizeMergeFieldsEventArgs)
            Dim mergeFieldNames As List(Of MergeFieldName) = New List(Of MergeFieldName)(e.MergeFieldsNames)
            mergeFieldNames.Remove(mergeFieldNames.Find(Function(mfn) Equals(mfn.Name.ToLower(), "password")))
            mergeFieldNames.ForEach(New Action(Of MergeFieldName)(AddressOf ChangeDisplayName))
            mergeFieldNames.Sort(New ReverseComparer())
            e.MergeFieldsNames = mergeFieldNames.ToArray()
        End Sub

        Private Shared Sub ChangeDisplayName(ByVal mfn As MergeFieldName)
            mfn.DisplayName += " (field)"
        End Sub
    End Class

    Public Class ReverseComparer
        Implements IComparer(Of MergeFieldName)

'#Region "IComparer<MergeFieldName> Members"
        Public Function Compare(ByVal mfn1 As MergeFieldName, ByVal mfn2 As MergeFieldName) As Integer Implements IComparer(Of MergeFieldName).Compare
            Return -String.Compare(mfn1.Name, mfn2.Name)
        End Function
'#End Region
    End Class

    Public Class Invoice

        Private idField As Integer

        Public Property Id As Integer
            Get
                Return idField
            End Get

            Set(ByVal value As Integer)
                idField = value
            End Set
        End Property

        Private descriptionField As String

        Public Property Description As String
            Get
                Return descriptionField
            End Get

            Set(ByVal value As String)
                descriptionField = value
            End Set
        End Property

        Private priceField As Decimal

        Public Property Price As Decimal
            Get
                Return priceField
            End Get

            Set(ByVal value As Decimal)
                priceField = value
            End Set
        End Property

        Private passwordField As Guid

        Public Property Password As Guid
            Get
                Return passwordField
            End Get

            Set(ByVal value As Guid)
                passwordField = value
            End Set
        End Property

        Public Sub New(ByVal id As Integer, ByVal description As String, ByVal price As Decimal, ByVal password As Guid)
            Me.Id = id
            Me.Description = description
            Me.Price = price
            Me.Password = password
        End Sub
    End Class
End Namespace
