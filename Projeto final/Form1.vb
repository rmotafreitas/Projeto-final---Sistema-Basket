Public Class Form

    Class ListNumSorter

        Implements IComparer
        Public Function CompareTo(ByVal o1 As Object,
                ByVal o2 As Object) As Integer _
                Implements System.Collections.IComparer.Compare
            Dim item1, item2 As ListViewItem
            item1 = CType(o1, ListViewItem)
            item2 = CType(o2, ListViewItem)
            Dim val1 = Convert.ToInt32(item1.SubItems(1).Text)
            Dim val2 = Convert.ToInt32(item2.SubItems(1).Text)

            Return val1 - val2
        End Function
    End Class


    Sub clear()
        TxtNum.Clear()
        TxtNome.Clear()
        TxtIdade.Clear()
        TxtPeso.Clear()
        TxtAltura.Clear()
        Select Case True
            Case RbtnAzul.Checked
                RbtnAzul.Checked = False
            Case RbtnCinzento.Checked
                RbtnCinzento.Checked = False
            Case RbtnLaranja.Checked
                RbtnLaranja.Checked = False
        End Select
    End Sub



    Private Function validarDados() As Boolean

        ' ====================== Verificação ====================== 
        Try
            If Convert.ToInt32(TxtNum.Text) <= 0 Then
                MessageBox.Show("Deve ser preenchido com um valor númerico maior que 0 (Número)!", "Warnig", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TxtNum.Focus()
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("Deve ser preenchido com um valor númerico (Número)!", "Warnig", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TxtNum.Focus()
            Return False
        End Try

        If TxtNome.Text = "" Then
            MessageBox.Show("Está faltando o prenchimento do seu nome!", "Warnig", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TxtNome.Focus()
            Return False
        End If

        Try
            If Convert.ToInt32(TxtIdade.Text) <= 0 Then
                MessageBox.Show("Deve ser preenchido com um valor númerico maior que 0 (Idade)!", "Warnig", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TxtIdade.Focus()
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("Deve ser preenchido com um valor númerico (Idade)!", "Warnig", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TxtIdade.Focus()
            Return False
        End Try

        Try
            If Convert.ToDouble(TxtPeso.Text) <= 0 Then
                MessageBox.Show("Deve ser preenchido com um valor númerico maior que 0 (Peso)!", "Warnig", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TxtPeso.Focus()
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("Deve ser preenchido com um valor númerico (Peso)!", "Warnig", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TxtIdade.Focus()
            Return False
        End Try

        Try
            If Convert.ToDouble(TxtAltura.Text) <= 0 Then
                MessageBox.Show("Deve ser preenchido com um valor númerico maior que 0 (Altura)!", "Warnig", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TxtAltura.Focus()
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("Deve ser preenchido com um valor númerico (Altura)!", "Warnig", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TxtAltura.Focus()
            Return False
        End Try

        'radio
        If RbtnAzul.Checked = False And RbtnCinzento.Checked = False And RbtnLaranja.Checked = False Then
            MessageBox.Show("Deve colocar a cor desejada!", "Warnig", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            GpbDados.Focus()
            Return False
        End If


        For Each _itm In lstvDados.Items
            If (lstvDados.SelectedItems.Count > 0) Then

                If (lstvDados.SelectedItems(0).Index <> _itm.Index And _itm.SubItems(1).Text.ToString.Equals(TxtNum.Text)) Then

                    MessageBox.Show("Não insira o mesmo número duas vezes!", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    TxtNum.Focus()
                    Return False

                End If

            End If
        Next

        If BtnAdd.Enabled = True Then
            For Each _itm In lstvDados.Items

                If (_itm.SubItems(1).Text).ToString.Equals(TxtNum.Text) Then

                    MessageBox.Show("Não insira o mesmo número duas vezes!", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    TxtNum.Focus()
                    Return False

                End If

            Next
        End If

        '=====================================

        Return True
    End Function

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click


        If (validarDados() = False) Then
            Return
        End If

        ' objeto item (listViewItem) para ser adicionado a lista de itens da listview
        ' Pois a ListView é um conjunto de listViewItem's



        ' Criar uma listViewItem - deve criar logo com o conteúdo da primeira coluna que se pretende
        ' deve ter a mesma estrutura da ListView ( as mesmas colunas )

        Dim li As ListViewItem = New ListViewItem("")

        'Dim li As ListViewItem = lstvDados.Items.Add("")

        li.SubItems.Add(TxtNum.Text)
        li.SubItems.Add(TxtNome.Text)
        li.SubItems.Add(TxtIdade.Text)
        li.SubItems.Add(TxtPeso.Text)
        li.SubItems.Add(TxtAltura.Text)
        li.SubItems.Add("")

        Select Case True
            Case RbtnAzul.Checked
                li.Tag = 0
            Case RbtnCinzento.Checked
                li.Tag = 1
            Case RbtnLaranja.Checked
                li.Tag = 2
        End Select

        lstvDados.Items.Add(li)


        clear()

        If lstvDados.Items.Count > 1 Then
            BtnOrdenar.Enabled = True
        End If

        TxtRegs.Text = lstvDados.Items.Count

        BtnReset.Enabled = True

    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles BtnClose.Click
        Close()
    End Sub

    Private Sub BtnOrdenar_Click(sender As Object, e As EventArgs) Handles BtnOrdenar.Click
        lstvDados.ListViewItemSorter = New ListNumSorter()
        lstvDados.Sort()

        BtnOrdenar.Enabled = False

        lstvDados.ListViewItemSorter = Nothing
        lstvDados.Sorting = SortOrder.None


    End Sub

    Private Sub lstvDados_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstvDados.SelectedIndexChanged
        If lstvDados.SelectedItems.Count <= 0 Then
            BtnAlterar.Enabled = False
            BtnRemover.Enabled = False
        End If


    End Sub

    Private Sub BtnAlterar_Click(sender As Object, e As EventArgs) Handles BtnAlterar.Click

        BtnAdd.Enabled = False
        BtnOrdenar.Enabled = False
        BtnRemover.Enabled = False
        BtnAlterar.Enabled = False

        'Get select itens and set to text boxs
        TxtNum.Text = lstvDados.SelectedItems(0).SubItems(1).Text
        TxtNome.Text = lstvDados.SelectedItems(0).SubItems(2).Text
        TxtIdade.Text = lstvDados.SelectedItems(0).SubItems(3).Text
        TxtPeso.Text = lstvDados.SelectedItems(0).SubItems(4).Text
        TxtAltura.Text = lstvDados.SelectedItems(0).SubItems(5).Text

        Select Case lstvDados.SelectedItems(0).Tag
            Case 0
                RbtnAzul.Checked = True
            Case 1
                RbtnCinzento.Checked = True
            Case 2
                RbtnLaranja.Checked = True
        End Select

        'Enabled
        BtnConfirmar.Enabled = True
        BtnCancela.Enabled = True

    End Sub

    Private Sub lstvDados_MouseClick(sender As Object, e As MouseEventArgs) Handles lstvDados.MouseClick
        If lstvDados.SelectedItems.Count > 0 Then
            BtnAlterar.Enabled = True
            BtnRemover.Enabled = True
        End If

        TxtQntsJogar.Text = lstvDados.CheckedItems.Count

    End Sub

    Private Sub BtnRemover_Click(sender As Object, e As EventArgs) Handles BtnRemover.Click
        If lstvDados.SelectedItems(0).Checked = True Then
            TxtQntsJogar.Text -= 1
        End If

        lstvDados.Items.Remove(lstvDados.SelectedItems(0))

        If lstvDados.Items.Count <= 1 Then
            BtnOrdenar.Enabled = False
            If lstvDados.Items.Count = 1 Then
                BtnReset.Enabled = True
            End If
        End If

        TxtRegs.Text = lstvDados.Items.Count
    End Sub

    Private Sub BtnCancela_Click(sender As Object, e As EventArgs) Handles BtnCancela.Click
        clear()
        BtnConfirmar.Enabled = False
        BtnCancela.Enabled = False
        BtnAdd.Enabled = True
    End Sub

    Private Sub BtnConfirmar_Click(sender As Object, e As EventArgs) Handles BtnConfirmar.Click

        If (validarDados() = False) Then
            Return
        End If

        lstvDados.SelectedItems(0).SubItems(1).Text = TxtNum.Text
        lstvDados.SelectedItems(0).SubItems(2).Text = TxtNome.Text
        lstvDados.SelectedItems(0).SubItems(3).Text = TxtIdade.Text
        lstvDados.SelectedItems(0).SubItems(4).Text = TxtPeso.Text
        lstvDados.SelectedItems(0).SubItems(5).Text = TxtAltura.Text

        Select Case True
            Case RbtnAzul.Checked
                lstvDados.SelectedItems(0).Tag = 0
            Case RbtnCinzento.Checked
                lstvDados.SelectedItems(0).Tag = 1
            Case RbtnLaranja.Checked
                lstvDados.SelectedItems(0).Tag = 2
        End Select

        clear()

        BtnConfirmar.Enabled = False
        BtnCancela.Enabled = False
        BtnAdd.Enabled = True
        If lstvDados.Items.Count > 1 Then
            BtnOrdenar.Enabled = True
        End If

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles BtnReset.Click
        If MessageBox.Show("Têm a certeza que quer limpar tudo?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            clear()
            BtnOrdenar.Enabled = False
            BtnReset.Enabled = False
            lstvDados.Items.Clear()
            TxtRegs.Text = lstvDados.Items.Count
            BtnAdd.Enabled = True
            TxtQntsJogar.Text = 0
        End If

    End Sub

    Private Sub lstvDados_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles lstvDados.ItemChecked
        TxtQntsJogar.Text = lstvDados.CheckedItems.Count
    End Sub

    Private Sub lstvDados_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs) Handles lstvDados.DrawColumnHeader
        e.DrawDefault = True
    End Sub

    Private Sub lstvDados_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs) Handles lstvDados.DrawSubItem
        Dim imageindex As Integer = DirectCast(e.Item.Tag, Integer)
        If e.ColumnIndex = 6 Then
            e.Graphics.DrawImage(ImglstCores.Images(imageindex), e.Bounds.X, e.Bounds.Y, e.Bounds.Height, e.Bounds.Height)
        Else
            e.DrawDefault = True
        End If
    End Sub

    Private Sub lstvDados_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lstvDados.ColumnWidthChanging
        e.NewWidth = Me.lstvDados.Columns(e.ColumnIndex).Width
        e.Cancel = True
    End Sub
End Class
