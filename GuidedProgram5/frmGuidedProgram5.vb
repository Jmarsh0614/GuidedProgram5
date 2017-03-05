'Program Name : Flooring Cost Calculator Application
'Author : Joshua Marshall
'Date: March 5th, 2017
'Purpose: THis Windows Application computes the estimated cost 
'         of flooring based on the number of square feet and the 
'         following cost per square foot: Tile - $4.49 per square foot; 
'         Carpet - $4.99 per square foot: Hardwood - $7.49 per square foot.

Option Strict On


Public Class frmGuidedProgram5
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        'The btnCaclulate event handler calculates the estimated cost of flooring
        'based on the square foootage and the flooring type

        'Declaration Section
        Dim decFootage As Decimal
        Dim decCostPerSquareFoot As Decimal
        Dim decCostEstimate As Decimal
        Dim decTileCost As Decimal = 4.49D
        Dim decCarpetCost As Decimal = 4.99D
        Dim decHardwoodCost As Decimal = 7.49D

        'Did user enter a numeric value?
        'If user enters a valued that is numeric, convert it to a decimal and then set decFootage equal to the value
        If IsNumeric(txtFootage.Text) Then
            decFootage = Convert.ToDecimal(txtFootage.Text)


            'Is Square Footage Greater than Zero?
            If decFootage > 0 Then
                'Determine cost per square foot
                If radTile.Checked Then
                    decCostPerSquareFoot = decTileCost
                ElseIf radCarpet.Checked Then
                    decCostPerSquareFoot = decCarpetCost
                ElseIf radHardwood.Checked Then
                    decCostPerSquareFoot = decHardwoodCost
                End If

                'Caclulate and display the cost Estimate
                decCostEstimate = decFootage * decCostPerSquareFoot
                lblCostEstimate.Text = decCostEstimate.ToString("C")
            Else
                'Display error message if user entered a negative value
                MsgBox("You entered " & decFootage.ToString() & ". Enter a Postive Number", , "Input Error")
                txtFootage.Text = ""
                txtFootage.Focus()


            End If
        Else
            'Display error message if user entered a Non-Numeric Value
            MsgBox("Enther the Square Footage of Flooring", , "Input Error")

        End If



    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtFootage.Clear()
        lblCostEstimate.Text = ""
        radTile.Checked = True
        radCarpet.Checked = False
        radHardwood.Checked = False
        txtFootage.Focus()
    End Sub

    Private Sub frmGuidedProgram5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtFootage.Focus()
        lblCostEstimate.Text = ""


    End Sub
End Class
