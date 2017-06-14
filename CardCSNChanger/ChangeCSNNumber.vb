Public Class ChangeCSNNumber

    Dim errormsg As String = "Error Message"
    Dim ACMeterCSN As String
    Dim MAcauPassCSN As Integer
    Dim CSNFunction As New Functions

    Private Sub BTExit_Click(sender As Object, e As EventArgs) Handles BTExit.Click
        Me.Close()

    End Sub

    Private Sub BTExit1_Click(sender As Object, e As EventArgs) Handles BTExit1.Click
        Me.Close()

    End Sub

    Private Sub BtClearCSN_Click(sender As Object, e As EventArgs) Handles BtClearCSN.Click
        Txtcsn.Clear()
        Txtcsn.Focus()
    End Sub

    Private Sub BtCopy_Click(sender As Object, e As EventArgs) Handles BtCopy.Click
        Clipboard.SetText(TxtBtams.Text)
    End Sub

    Private Sub BtChangeCSN1_Click(sender As Object, e As EventArgs) Handles BtChangeCSN1.Click
        Try


            Dim oldCSN As String = Txtcsn.Text

            If Txtcsn.Text = String.Empty Then
                MessageBox.Show(errormsg, "The number is empty, please try again!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim deci = CULng("&H" & oldCSN.Trim)

            TxtBtams.Text = deci
            TxtoldCSN.Text = CSNFunction.reorderstring(oldCSN.Trim)

            TxtmutilCSN.Text &= " AC Meter car CSN number is:  " & Txtcsn.Text & "   AC Meter car number is:  " & deci & vbCrLf


        Catch
            MessageBox.Show(errormsg, "The number is error, please try again!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Txtcsn.Focus()
            Txtcsn.Text = vbNullString
        End Try
    End Sub

    Private Sub BtClearAll1_Click(sender As Object, e As EventArgs) Handles BtClearAll1.Click
        Txtcsn.Clear()
        TxtBtams.Clear()
        TxtoldCSN.Clear()
        TxtmutilCSN.Clear()
        Txtcsn.Focus()
    End Sub

    Private Sub BtChangeCSN_Click(sender As Object, e As EventArgs) Handles BtChangeCSN.Click
        Try


            Dim oldCSN As String = Txtcsn.Text

            If Txtcsn.Text = String.Empty Then
                MessageBox.Show(errormsg, "The number is empty, please try again!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim deci = CULng("&H" & oldCSN.Trim)

            TxtBtams.Text = deci
            TxtoldCSN.Text = CSNFunction.reorderstring(oldCSN.Trim)

            TxtmutilCSN.Text &= " AC Meter car CSN number is:  " & Txtcsn.Text & "   AC Meter car number is:  " & deci & vbCrLf


        Catch
            MessageBox.Show(errormsg, "The number is error, please try again!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Txtcsn.Focus()
            Txtcsn.Text = vbNullString
        End Try
    End Sub

    Private Sub BtClearAll_Click(sender As Object, e As EventArgs) Handles BtClearAll.Click
        Txtcsn.Clear()
        TxtBtams.Clear()
        TxtoldCSN.Clear()
        TxtmutilCSN.Clear()
        Txtcsn.Focus()
    End Sub

    Private Sub Txtnewcsn_keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Txtcsn.KeyPress
        Try
            Dim strlimit As String
            strlimit = "0123456789abcdef"
            Dim keychar As Char = e.KeyChar
            If InStr(strlimit, keychar) <> 0 Or e.KeyChar = Microsoft.VisualBasic.ChrW(8) Then
                e.Handled = False
            Else
                e.Handled = True
            End If

        Catch
            MessageBox.Show(errormsg, "The number is error, please try again!!", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub


    Private Sub ChangeCSNNumber_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Txtcsn.Focus()
    End Sub

    Private Sub BTClearMP2_Click(sender As Object, e As EventArgs) Handles BTClearMP2.Click
        Me.Close()
    End Sub

    Private Sub BTClearMP_Click(sender As Object, e As EventArgs) Handles BTClearMP.Click
        Me.Close()
    End Sub

    Private Sub BTClearAllMP_Click(sender As Object, e As EventArgs) Handles BTClearAllMP.Click
        TxtCSNMP.Clear()
        TxtBTAMSMP.Clear()
        TxtOldCSNMP.Clear()
        TxtMutilCSNMP.Clear()
        TxtCSNMP.Focus()
    End Sub

    Private Sub BTClearAllMP2_Click(sender As Object, e As EventArgs) Handles BTClearAllMP2.Click
        TxtCSNMP.Clear()
        TxtBTAMSMP.Clear()
        TxtOldCSNMP.Clear()
        TxtMutilCSNMP.Clear()
        TxtCSNMP.Focus()
    End Sub

    Private Sub BTChangeCSNMP_Click(sender As Object, e As EventArgs) Handles BTChangeCSNMP.Click
        Try


            Dim oldCSN As String = TxtCSNMP.Text
            Dim ReorderCSN As String

            If TxtCSNMP.Text = String.Empty Then
                MessageBox.Show(errormsg, "The number is empty, please try again!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ReorderCSN = CSNFunction.reorderstring(oldCSN.Trim)
            Dim deci = CULng("&H" & ReorderCSN.Trim)

            TxtBTAMSMP.Text = deci
            TxtOldCSNMP.Text = ReorderCSN

            TxtMutilCSNMP.Text &= " Macau Pass car CSN number is:  " & TxtCSNMP.Text & "   Macau Pass BTAMS number is:  " & deci & vbCrLf


        Catch
            MessageBox.Show(errormsg, "The number is error, please try again!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TxtCSNMP.Focus()
            TxtCSNMP.Text = vbNullString
        End Try
    End Sub

    Private Sub BTChangeCSNMP2_Click(sender As Object, e As EventArgs) Handles BTChangeCSNMP2.Click
        Try


            Dim oldCSN As String = TxtCSNMP.Text
            Dim ReorderCSN As String

            If TxtCSNMP.Text = String.Empty Then
                MessageBox.Show(errormsg, "The number is empty, please try again!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ReorderCSN = CSNFunction.reorderstring(oldCSN.Trim)
            Dim deci = CULng("&H" & ReorderCSN.Trim)

            TxtBTAMSMP.Text = deci
            TxtOldCSNMP.Text = ReorderCSN

            TxtMutilCSNMP.Text &= " Macau Pass car CSN number is:  " & TxtCSNMP.Text & "   Macau Pass BTAMS number is:  " & deci & vbCrLf


        Catch
            MessageBox.Show(errormsg, "The number is error, please try again!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TxtCSNMP.Focus()
            TxtCSNMP.Text = vbNullString
        End Try
    End Sub

    Private Sub BTClearcsnMP_Click(sender As Object, e As EventArgs) Handles BTClearcsnMP.Click
        TxtCSNMP.Clear()
    End Sub

    Private Sub TxtcsnMP_keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCSNMP.KeyPress
        Try
            Dim strlimit As String
            strlimit = "0123456789abcdef"
            Dim keychar As Char = e.KeyChar
            If InStr(strlimit, keychar) <> 0 Or e.KeyChar = Microsoft.VisualBasic.ChrW(8) Then
                e.Handled = False
            Else
                e.Handled = True
            End If

        Catch
            MessageBox.Show(errormsg, "The number is error, please try again!!", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub BTCopyMP_Click(sender As Object, e As EventArgs) Handles BTCopyMP.Click
        Clipboard.SetText(TxtBTAMSMP.Text)
    End Sub
End Class
