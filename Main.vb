Imports System.IO

Public Class Delivergrid
    Inherits System.Windows.Forms.Form

#Region " Variables "
    '-- Location Variables
    Dim LeftStart As Integer = 16
    Dim LeftIncrement As Integer = 32
    Dim TopStart As Integer = 16
    Dim TopIncrement As Integer = 32

    Dim LegendLeft As Integer = 8
    Dim LegendTop As Integer = 16
    Dim LegendHIncrement As Integer = 80
    Dim LegendVIncrement As Integer = 40

    Dim InventoryLeft As Integer = 20
    Dim InventoryTop As Integer = 20
    Dim InventoryHIncrement As Integer = 40
    Dim InventoryVIncrement As Integer = 40
    '--
    Dim Points As New ArrayList
    Dim BigFont As Integer = 12
    Dim SmallFont As Integer = 10
    '--
    Dim ColorArray(ColorCount) As Color
    Dim DirectionArray(DirectionCount) As String
    '--
    Dim CurrentLevel As New Level
    '--
    Dim LastClickedX As Integer = 0
    Dim LastClickedY As Integer = 0
    Dim LastClickedLegend As Integer = 0
    Dim LastClickedInvX As Integer = 0
    Dim LastClickedInvY As Integer = 0
    '--
    Dim NumberingGoals As Boolean = False
    Dim GoalNumber As Integer = 0
    Dim TotalGoals As Integer = 0
    Dim GoalList() As Pair
    '--
    Dim TotalPathLength As Double
    Dim PathStack As New Stack
    Dim PathList As New ArrayList
    '--
    Dim Permutations As New ArrayList
    '--
    Dim StepSpeed As Double = 1.0
    Dim StepSpeedDiagonal As Double = 1.5
    '--
    Dim init As Boolean = False

#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents gbLegend As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtLevelName As System.Windows.Forms.TextBox
    Friend WithEvents lblLevel As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents rbPlay As System.Windows.Forms.RadioButton
    Friend WithEvents rbEdit As System.Windows.Forms.RadioButton
    Friend WithEvents btSave As System.Windows.Forms.Button
    Friend WithEvents txtCurrentSolve As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents numLevel As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtStartX As System.Windows.Forms.TextBox
    Friend WithEvents txtStartY As System.Windows.Forms.TextBox
    Friend WithEvents btNumberGoals As System.Windows.Forms.Button
    Friend WithEvents btSolve As System.Windows.Forms.Button
    Friend WithEvents btViewPath As System.Windows.Forms.Button
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents btLoad As System.Windows.Forms.Button
    Friend WithEvents btClear As System.Windows.Forms.Button
    Friend WithEvents numTarget As System.Windows.Forms.NumericUpDown
    Friend WithEvents btReset As System.Windows.Forms.Button
    Friend WithEvents gbInventory As System.Windows.Forms.GroupBox
    Friend WithEvents gbMain As System.Windows.Forms.GroupBox
    Friend WithEvents btSolveTSP As System.Windows.Forms.Button
    Friend WithEvents cbPresetOrder As System.Windows.Forms.CheckBox
    Friend WithEvents btRules As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Delivergrid))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.gbLegend = New System.Windows.Forms.GroupBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btClear = New System.Windows.Forms.Button()
        Me.txtLevelName = New System.Windows.Forms.TextBox()
        Me.numLevel = New System.Windows.Forms.NumericUpDown()
        Me.lblLevel = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.rbPlay = New System.Windows.Forms.RadioButton()
        Me.rbEdit = New System.Windows.Forms.RadioButton()
        Me.btSave = New System.Windows.Forms.Button()
        Me.gbInventory = New System.Windows.Forms.GroupBox()
        Me.btSolve = New System.Windows.Forms.Button()
        Me.txtCurrentSolve = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.btNumberGoals = New System.Windows.Forms.Button()
        Me.txtStartX = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtStartY = New System.Windows.Forms.TextBox()
        Me.btViewPath = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.btLoad = New System.Windows.Forms.Button()
        Me.numTarget = New System.Windows.Forms.NumericUpDown()
        Me.btRules = New System.Windows.Forms.Button()
        Me.btReset = New System.Windows.Forms.Button()
        Me.gbMain = New System.Windows.Forms.GroupBox()
        Me.cbPresetOrder = New System.Windows.Forms.CheckBox()
        Me.btSolveTSP = New System.Windows.Forms.Button()
        Me.gbLegend.SuspendLayout()
        CType(Me.numLevel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numTarget, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.ControlDark
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(384, 384)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Board"
        Me.Label1.Visible = False
        '
        'gbLegend
        '
        Me.gbLegend.BackColor = System.Drawing.SystemColors.ControlDark
        Me.gbLegend.Controls.Add(Me.Label9)
        Me.gbLegend.Controls.Add(Me.Label8)
        Me.gbLegend.Controls.Add(Me.Label7)
        Me.gbLegend.Controls.Add(Me.Label6)
        Me.gbLegend.Controls.Add(Me.Label5)
        Me.gbLegend.Controls.Add(Me.Label4)
        Me.gbLegend.Controls.Add(Me.Label3)
        Me.gbLegend.Controls.Add(Me.Label2)
        Me.gbLegend.Controls.Add(Me.btClear)
        Me.gbLegend.Location = New System.Drawing.Point(16, 408)
        Me.gbLegend.Name = "gbLegend"
        Me.gbLegend.Size = New System.Drawing.Size(384, 96)
        Me.gbLegend.TabIndex = 2
        Me.gbLegend.TabStop = False
        Me.gbLegend.Text = "Legend"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(280, 64)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(48, 16)
        Me.Label9.TabIndex = 14
        Me.Label9.Text = "Teleport"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(200, 64)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(32, 16)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Sink"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(120, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 16)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "Fast"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(40, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(48, 16)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Jump"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(280, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(32, 16)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Goal"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(200, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 16)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Solid"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(120, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 16)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Wall"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(40, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Regular"
        '
        'btClear
        '
        Me.btClear.BackColor = System.Drawing.Color.LightCoral
        Me.btClear.Location = New System.Drawing.Point(328, 32)
        Me.btClear.Name = "btClear"
        Me.btClear.Size = New System.Drawing.Size(40, 40)
        Me.btClear.TabIndex = 22
        Me.btClear.Text = "Clear All"
        Me.btClear.UseVisualStyleBackColor = False
        '
        'txtLevelName
        '
        Me.txtLevelName.Location = New System.Drawing.Point(8, 40)
        Me.txtLevelName.Name = "txtLevelName"
        Me.txtLevelName.Size = New System.Drawing.Size(176, 20)
        Me.txtLevelName.TabIndex = 3
        '
        'numLevel
        '
        Me.numLevel.Location = New System.Drawing.Point(280, 104)
        Me.numLevel.Name = "numLevel"
        Me.numLevel.Size = New System.Drawing.Size(80, 20)
        Me.numLevel.TabIndex = 4
        Me.numLevel.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.numLevel.Visible = False
        '
        'lblLevel
        '
        Me.lblLevel.Location = New System.Drawing.Point(280, 88)
        Me.lblLevel.Name = "lblLevel"
        Me.lblLevel.Size = New System.Drawing.Size(48, 16)
        Me.lblLevel.TabIndex = 5
        Me.lblLevel.Text = "Level #"
        Me.lblLevel.Visible = False
        '
        'Label10
        '
        Me.Label10.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label10.Location = New System.Drawing.Point(8, 24)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 16)
        Me.Label10.TabIndex = 6
        Me.Label10.Text = "Level Name"
        '
        'rbPlay
        '
        Me.rbPlay.Location = New System.Drawing.Point(8, 152)
        Me.rbPlay.Name = "rbPlay"
        Me.rbPlay.Size = New System.Drawing.Size(48, 16)
        Me.rbPlay.TabIndex = 7
        Me.rbPlay.Text = "Play"
        '
        'rbEdit
        '
        Me.rbEdit.Checked = True
        Me.rbEdit.Location = New System.Drawing.Point(56, 152)
        Me.rbEdit.Name = "rbEdit"
        Me.rbEdit.Size = New System.Drawing.Size(48, 16)
        Me.rbEdit.TabIndex = 8
        Me.rbEdit.TabStop = True
        Me.rbEdit.Text = "Edit"
        '
        'btSave
        '
        Me.btSave.BackColor = System.Drawing.Color.CornflowerBlue
        Me.btSave.Location = New System.Drawing.Point(112, 72)
        Me.btSave.Name = "btSave"
        Me.btSave.Size = New System.Drawing.Size(72, 24)
        Me.btSave.TabIndex = 9
        Me.btSave.Text = "Save"
        Me.btSave.UseVisualStyleBackColor = False
        '
        'gbInventory
        '
        Me.gbInventory.BackColor = System.Drawing.SystemColors.ControlDark
        Me.gbInventory.Location = New System.Drawing.Point(416, 256)
        Me.gbInventory.Name = "gbInventory"
        Me.gbInventory.Size = New System.Drawing.Size(192, 144)
        Me.gbInventory.TabIndex = 10
        Me.gbInventory.TabStop = False
        Me.gbInventory.Text = "Inventory"
        '
        'btSolve
        '
        Me.btSolve.BackColor = System.Drawing.Color.CornflowerBlue
        Me.btSolve.Location = New System.Drawing.Point(416, 432)
        Me.btSolve.Name = "btSolve"
        Me.btSolve.Size = New System.Drawing.Size(120, 24)
        Me.btSolve.TabIndex = 11
        Me.btSolve.Text = "Solve Current Order"
        Me.btSolve.UseVisualStyleBackColor = False
        '
        'txtCurrentSolve
        '
        Me.txtCurrentSolve.Location = New System.Drawing.Point(544, 432)
        Me.txtCurrentSolve.Name = "txtCurrentSolve"
        Me.txtCurrentSolve.Size = New System.Drawing.Size(64, 20)
        Me.txtCurrentSolve.TabIndex = 12
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 64)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(88, 16)
        Me.Label11.TabIndex = 13
        Me.Label11.Text = "Target Distance"
        '
        'btNumberGoals
        '
        Me.btNumberGoals.BackColor = System.Drawing.Color.Gold
        Me.btNumberGoals.Location = New System.Drawing.Point(416, 216)
        Me.btNumberGoals.Name = "btNumberGoals"
        Me.btNumberGoals.Size = New System.Drawing.Size(88, 24)
        Me.btNumberGoals.TabIndex = 15
        Me.btNumberGoals.Text = "Order Goals"
        Me.btNumberGoals.UseVisualStyleBackColor = False
        '
        'txtStartX
        '
        Me.txtStartX.Location = New System.Drawing.Point(280, 160)
        Me.txtStartX.Name = "txtStartX"
        Me.txtStartX.ReadOnly = True
        Me.txtStartX.Size = New System.Drawing.Size(48, 20)
        Me.txtStartX.TabIndex = 17
        Me.txtStartX.Visible = False
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(280, 144)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(80, 16)
        Me.Label12.TabIndex = 16
        Me.Label12.Text = "Start Location"
        Me.Label12.Visible = False
        '
        'txtStartY
        '
        Me.txtStartY.Location = New System.Drawing.Point(328, 160)
        Me.txtStartY.Name = "txtStartY"
        Me.txtStartY.ReadOnly = True
        Me.txtStartY.Size = New System.Drawing.Size(48, 20)
        Me.txtStartY.TabIndex = 18
        Me.txtStartY.Visible = False
        '
        'btViewPath
        '
        Me.btViewPath.BackColor = System.Drawing.Color.CornflowerBlue
        Me.btViewPath.Location = New System.Drawing.Point(544, 464)
        Me.btViewPath.Name = "btViewPath"
        Me.btViewPath.Size = New System.Drawing.Size(64, 24)
        Me.btViewPath.TabIndex = 19
        Me.btViewPath.Text = "View Path"
        Me.btViewPath.UseVisualStyleBackColor = False
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(544, 416)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(48, 16)
        Me.Label13.TabIndex = 20
        Me.Label13.Text = "Distance"
        '
        'btLoad
        '
        Me.btLoad.BackColor = System.Drawing.Color.CornflowerBlue
        Me.btLoad.Location = New System.Drawing.Point(112, 104)
        Me.btLoad.Name = "btLoad"
        Me.btLoad.Size = New System.Drawing.Size(72, 24)
        Me.btLoad.TabIndex = 21
        Me.btLoad.Text = "Load"
        Me.btLoad.UseVisualStyleBackColor = False
        '
        'numTarget
        '
        Me.numTarget.Location = New System.Drawing.Point(8, 80)
        Me.numTarget.Name = "numTarget"
        Me.numTarget.Size = New System.Drawing.Size(80, 20)
        Me.numTarget.TabIndex = 23
        Me.numTarget.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btRules
        '
        Me.btRules.BackColor = System.Drawing.Color.Pink
        Me.btRules.Location = New System.Drawing.Point(112, 144)
        Me.btRules.Name = "btRules"
        Me.btRules.Size = New System.Drawing.Size(72, 24)
        Me.btRules.TabIndex = 24
        Me.btRules.Text = "Rules"
        Me.btRules.UseVisualStyleBackColor = False
        '
        'btReset
        '
        Me.btReset.BackColor = System.Drawing.Color.LightCoral
        Me.btReset.Location = New System.Drawing.Point(520, 216)
        Me.btReset.Name = "btReset"
        Me.btReset.Size = New System.Drawing.Size(88, 24)
        Me.btReset.TabIndex = 25
        Me.btReset.Text = "Reset Level"
        Me.btReset.UseVisualStyleBackColor = False
        '
        'gbMain
        '
        Me.gbMain.BackColor = System.Drawing.SystemColors.ControlDark
        Me.gbMain.Controls.Add(Me.rbEdit)
        Me.gbMain.Controls.Add(Me.rbPlay)
        Me.gbMain.Controls.Add(Me.Label10)
        Me.gbMain.Controls.Add(Me.Label11)
        Me.gbMain.Controls.Add(Me.txtLevelName)
        Me.gbMain.Controls.Add(Me.btLoad)
        Me.gbMain.Controls.Add(Me.numTarget)
        Me.gbMain.Controls.Add(Me.btSave)
        Me.gbMain.Controls.Add(Me.btRules)
        Me.gbMain.Controls.Add(Me.cbPresetOrder)
        Me.gbMain.Location = New System.Drawing.Point(416, 16)
        Me.gbMain.Name = "gbMain"
        Me.gbMain.Size = New System.Drawing.Size(192, 184)
        Me.gbMain.TabIndex = 26
        Me.gbMain.TabStop = False
        Me.gbMain.Text = "Overview"
        '
        'cbPresetOrder
        '
        Me.cbPresetOrder.Location = New System.Drawing.Point(8, 104)
        Me.cbPresetOrder.Name = "cbPresetOrder"
        Me.cbPresetOrder.Size = New System.Drawing.Size(88, 24)
        Me.cbPresetOrder.TabIndex = 0
        Me.cbPresetOrder.Text = "Preset Order"
        '
        'btSolveTSP
        '
        Me.btSolveTSP.BackColor = System.Drawing.Color.CornflowerBlue
        Me.btSolveTSP.Location = New System.Drawing.Point(416, 464)
        Me.btSolveTSP.Name = "btSolveTSP"
        Me.btSolveTSP.Size = New System.Drawing.Size(120, 24)
        Me.btSolveTSP.TabIndex = 27
        Me.btSolveTSP.Text = "Solve For Any Order"
        Me.btSolveTSP.UseVisualStyleBackColor = False
        '
        'Delivergrid
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ClientSize = New System.Drawing.Size(626, 512)
        Me.Controls.Add(Me.gbMain)
        Me.Controls.Add(Me.btReset)
        Me.Controls.Add(Me.txtStartY)
        Me.Controls.Add(Me.txtStartX)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.btNumberGoals)
        Me.Controls.Add(Me.gbInventory)
        Me.Controls.Add(Me.lblLevel)
        Me.Controls.Add(Me.numLevel)
        Me.Controls.Add(Me.gbLegend)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btSolveTSP)
        Me.Controls.Add(Me.txtCurrentSolve)
        Me.Controls.Add(Me.btSolve)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.btViewPath)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Delivergrid"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Delivergrid"
        Me.gbLegend.ResumeLayout(False)
        CType(Me.numLevel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numTarget, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbMain.ResumeLayout(False)
        Me.gbMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub


#Region " Initialization "
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        StartGame()
    End Sub

    Sub StartGame()
        'Dim Introduction As New Intro
        'If Introduction.ShowDialog() <> DialogResult.OK Then
        '    Me.DialogResult = DialogResult.No
        '    Me.Close()
        'End If

        FillColorArray()
        FillDirectionArray()
        CreateLegend()
        CreateInventory()
        UpdateInventory()
        CreateBoxes()
        UpdateGrid()
        SaveBackup()
        init = True
    End Sub

    Sub FillColorArray()
        ColorArray(TerrainRegular) = ColorRegular
        ColorArray(TerrainWall) = ColorWall
        ColorArray(TerrainUnbreakable) = ColorUnbreakable
        ColorArray(TerrainGoal) = ColorGoal
        ColorArray(TerrainJump) = ColorJump
        ColorArray(TerrainFast) = ColorFast
        ColorArray(TerrainSink) = ColorSink
        ColorArray(TerrainTeleport) = ColorTeleport
    End Sub

    Sub FillDirectionArray()
        DirectionArray(DirectionNull) = LetterNull
        DirectionArray(DirectionN) = LetterN
        DirectionArray(DirectionE) = LetterE
        DirectionArray(DirectionS) = LetterS
        DirectionArray(DirectionW) = LetterW
        DirectionArray(DirectionV) = LetterV
        DirectionArray(DirectionH) = LetterH
        DirectionArray(DirectionA) = LetterA
    End Sub

    Sub StartNewGame()
        Me.DialogResult = DialogResult.Yes
        Me.Close()
    End Sub

    Sub CreateLegend()
        Dim i As Integer
        Dim CurrentX As Integer = LegendLeft
        Dim CurrentY As Integer = LegendTop
        For i = 0 To ColorCount - 1
            '--Create legendinfo
            LegendInfo(i) = New TerrainSquare(i, 0)
            Dim NewLegend As New Label
            NewLegend.Parent = gbLegend
            NewLegend.BorderStyle = BorderStyle.FixedSingle
            NewLegend.Location = New System.Drawing.Point(gbLegend.Location.X + CurrentX, gbLegend.Location.Y + CurrentY)
            NewLegend.BringToFront()
            If i = 3 Then
                CurrentX = LegendLeft
                CurrentY += LegendVIncrement
            Else
                CurrentX += LegendHIncrement
            End If

            NewLegend.Name = "NewBox"
            NewLegend.Tag = i
            NewLegend.Size = New System.Drawing.Size(LeftIncrement, TopIncrement)
            NewLegend.TabStop = False
            NewLegend.TextAlign = ContentAlignment.MiddleCenter
            NewLegend.BackColor = ColorArray(i)
            'NewLegend.Appearance.FontData.SizeInPoints = SmallFont
            'NewLegend.Appearance.BackColor = Color.White
            'NewLegend.Appearance.BackColor2 = ColorArray(i)
            'NewLegend.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Rectangular
            'NewLegend.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            'NewLegend.Appearance.TextVAlign = Infragistics.Win.VAlign.Middle
            NewLegend.Text = ""
            AddHandler NewLegend.MouseUp, AddressOf LabelMouseUpLegend
            TheLegend(i) = NewLegend
            Me.Controls.Add(NewLegend)
            NewLegend.BringToFront()
        Next
    End Sub

    Private Sub LabelMouseUpLegend(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim TheLegendBox As Label = CType(sender, Label)
        Dim i As Integer = TheLegendBox.Tag
        TheLegend(LastClickedLegend).BorderStyle = BorderStyle.FixedSingle
        LastClickedLegend = i
        TheLegend(LastClickedLegend).BorderStyle = BorderStyle.Fixed3D
        '--
        If e.Button = MouseButtons.Left Then
            '
        ElseIf e.Button = MouseButtons.Right Then
            '-- Currently used for view mode
        End If

    End Sub

    Sub CreateInventory()
        Dim i, j As Integer
        Dim CurrentX As Integer = InventoryLeft
        Dim CurrentY As Integer = InventoryTop
        For j = 0 To InventoryHeight
            For i = 0 To InventoryWidth
                '--Create Inventoryinfo
                InventoryInfo(i, j) = New TerrainSquare(i, j, -1, 0)

                Dim NewInventory As New Label
                NewInventory.BorderStyle = BorderStyle.FixedSingle
                NewInventory.Location = New System.Drawing.Point(gbInventory.Location.X + CurrentX, gbInventory.Location.Y + CurrentY)
                NewInventory.BringToFront()
                CurrentX += InventoryHIncrement
                NewInventory.Name = "NewInventory"
                NewInventory.Tag = i & "," & j
                NewInventory.Size = New System.Drawing.Size(LeftIncrement, TopIncrement)
                NewInventory.TabStop = False
                'NewInventory.Appearance.FontData.SizeInPoints = SmallFont
                'NewInventory.BackColor = Color.White;
                If InventoryInfo(i, j).TerrainType > 0 Then
                    NewInventory.BackColor = ColorArray(InventoryInfo(i, j).TerrainType)
                Else
                    NewInventory.BackColor = ColorNull
                End If
                NewInventory.TextAlign = ContentAlignment.MiddleCenter

                'NewInventory.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Rectangular
                'NewInventory.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
                'NewInventory.Appearance.TextVAlign = Infragistics.Win.VAlign.Middle
                'NewInventory.Text = "0"
                'AddHandler NewInventory.Click, AddressOf LabelClick
                'AddHandler NewInventory.MouseEnter, AddressOf LabelMouseOver
                AddHandler NewInventory.MouseUp, AddressOf LabelMouseUpInventory
                TheInventory(i, j) = NewInventory
                Me.Controls.Add(NewInventory)
                NewInventory.BringToFront()
            Next
            CurrentX = InventoryLeft
            CurrentY += InventoryVIncrement
        Next
    End Sub

    Private Sub LabelMouseUpInventory(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        'Dim TheInventory As Infragistics.Win.Misc.UltraButton = CType(sender, Infragistics.Win.Misc.UltraButton)
        Dim TheItem As Label = CType(sender, Label)
        Dim X, Y As Integer
        GetCoords(X, Y, TheItem.Tag)
        TheInventory(LastClickedInvX, LastClickedInvY).BorderStyle = BorderStyle.FixedSingle
        LastClickedInvX = X
        LastClickedInvY = Y
        TheInventory(LastClickedInvX, LastClickedInvY).BorderStyle = BorderStyle.Fixed3D
        '--
        If e.Button = MouseButtons.Left Then
            InventoryClicked(True)
        ElseIf e.Button = MouseButtons.Right Then
            InventoryClicked(False)
        End If

    End Sub

    Sub CreateBoxes()
        Dim i, j As Integer
        Dim CurrentX As Integer = LeftStart
        Dim CurrentY As Integer = TopStart
        For j = 0 To GridHeight
            For i = 0 To GridWidth
                '--Create boxinfo
                BoxInfo(i, j) = New TerrainSquare(i, j)

                Dim NewBox As New Label
                NewBox.BorderStyle = BorderStyle.FixedSingle
                NewBox.Location = New System.Drawing.Point(CurrentX, CurrentY)
                CurrentX += LeftIncrement
                NewBox.Name = "NewBox"
                NewBox.Tag = i & "," & j
                'NewBox.ContextMenu = ctxTerrain 'Return this if context menu is used
                NewBox.Size = New System.Drawing.Size(LeftIncrement, TopIncrement)
                NewBox.TabStop = False

                NewBox.TextAlign = ContentAlignment.MiddleCenter
                NewBox.BackColor = ColorArray(BoxInfo(i, j).TerrainType)

                'NewBox.Appearance.FontData.SizeInPoints = SmallFont
                'NewBox.Appearance.BackColor = Color.White
                'NewBox.Appearance.BackColor2 = ColorArray(BoxInfo(i, j).TerrainType)
                'NewBox.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Rectangular
                'NewBox.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
                'NewBox.Appearance.TextVAlign = Infragistics.Win.VAlign.Middle
                'NewBox.Text = "0"
                'AddHandler NewBox.Click, AddressOf LabelClick
                'AddHandler NewBox.MouseEnter, AddressOf LabelMouseOver
                AddHandler NewBox.MouseUp, AddressOf LabelMouseUp
                TheBoxes(i, j) = NewBox
                Me.Controls.Add(NewBox)
            Next
            CurrentX = LeftStart
            CurrentY += TopIncrement
        Next
    End Sub

    Private Sub LabelMouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        'Dim TheBox As Infragistics.Win.Misc.UltraButton = CType(sender, Infragistics.Win.Misc.UltraButton)
        Dim TheBox As Label = CType(sender, Label)
        Dim X, Y As Integer
        GetCoords(X, Y, TheBox.Tag)
        TheBoxes(LastClickedX, LastClickedY).BorderStyle = BorderStyle.FixedSingle
        LastClickedX = X
        LastClickedY = Y
        TheBoxes(LastClickedX, LastClickedY).BorderStyle = BorderStyle.Fixed3D
        '--
        If e.Button = MouseButtons.Left Then
            BoxClicked(True)
        ElseIf e.Button = MouseButtons.Right Then
            BoxClicked(False)
        End If

    End Sub

    Sub GetCoords(ByRef TheX As Integer, ByRef TheY As Integer, ByVal TheTag As String)
        Try
            Dim TheSections(1) As String
            TheSections = TheTag.Split(",")
            TheX = CType(TheSections(0), Integer)
            TheY = CType(TheSections(1), Integer)
        Catch ex As Exception
            MsgBox("Error getting coords.", , "Error")
        End Try

    End Sub
#End Region

#Region " Board "
    Sub BoxClicked(ByVal LeftClick As Boolean)
        '--Number the goals
        If NumberingGoals And BoxInfo(LastClickedX, LastClickedY).TerrainType = TerrainGoal And BoxInfo(LastClickedX, LastClickedY).Value <= 0 Then
            GoalNumber += 1
            BoxInfo(LastClickedX, LastClickedY).Value = GoalNumber
            If GoalNumber > TotalGoals Then
                NumberingGoals = False
            End If
        End If

        '--Use inventory item
        If rbPlay.Checked And BoxInfo(LastClickedX, LastClickedY).TerrainType <> TerrainGoal Then
            Dim InvTerrain As Integer = InventoryInfo(LastClickedInvX, LastClickedInvY).TerrainType
            Dim InvValue As Integer = InventoryInfo(LastClickedInvX, LastClickedInvY).Value
            If InvTerrain <> TerrainNull Then
                '--Can't use null item
                Dim SelectedTerrain As Integer = BoxInfo(LastClickedX, LastClickedY).TerrainType
                Dim SelectedValue As Integer = BoxInfo(LastClickedX, LastClickedY).Value
                If Not (InvTerrain = SelectedTerrain And InvValue = SelectedValue) Then
                    '--Prevent accidently using item redundantly
                    Dim Used As Boolean = False
                    If SelectedTerrain = TerrainRegular And SelectedValue < 1 Then
                        '--Place on free area
                        BoxInfo(LastClickedX, LastClickedY) = New TerrainSquare(InventoryInfo(LastClickedInvX, LastClickedInvY))
                        Used = True
                    End If
                    If (SelectedTerrain <> TerrainRegular Or SelectedValue > 0) And SelectedTerrain <> TerrainUnbreakable And InvTerrain = TerrainRegular Then
                        '--Clear location
                        BoxInfo(LastClickedX, LastClickedY) = New TerrainSquare(InventoryInfo(LastClickedInvX, LastClickedInvY))
                        Used = True
                    End If

                    If Used Then '--Clear item from inventory
                        InventoryInfo(LastClickedInvX, LastClickedInvY) = New TerrainSquare(LastClickedInvX, LastClickedInvY, TerrainNull, -1)
                    End If
                    UpdateInventory()
                End If
            End If
        End If

        '--Edit the board in edit mode
        If rbEdit.Checked Then
            If BoxInfo(LastClickedX, LastClickedY).TerrainType = LastClickedLegend Then
                If LeftClick Then
                    BoxInfo(LastClickedX, LastClickedY).Value += 1
                    If BoxInfo(LastClickedX, LastClickedY).TerrainType = TerrainRegular And BoxInfo(LastClickedX, LastClickedY).Value >= DirectionCount Then
                        '--Rotate letters off back
                        BoxInfo(LastClickedX, LastClickedY).Value = 0
                    End If
                ElseIf BoxInfo(LastClickedX, LastClickedY).Value > 0 Then
                    BoxInfo(LastClickedX, LastClickedY).Value -= 1
                ElseIf BoxInfo(LastClickedX, LastClickedY).TerrainType = TerrainRegular Then
                    '--Rotate letters off front
                    BoxInfo(LastClickedX, LastClickedY).Value = DirectionCount - 1
                End If
            Else
                BoxInfo(LastClickedX, LastClickedY).Value = 0
                BoxInfo(LastClickedX, LastClickedY).TerrainType = LastClickedLegend
            End If
        Else

        End If

        UpdateGrid()
    End Sub

    Sub InventoryClicked(ByVal LeftClick As Boolean)

        If rbEdit.Checked Then
            If InventoryInfo(LastClickedInvX, LastClickedInvY).TerrainType = LastClickedLegend Then
                If LeftClick Then
                    InventoryInfo(LastClickedInvX, LastClickedInvY).Value += 1
                    If InventoryInfo(LastClickedInvX, LastClickedInvY).TerrainType = TerrainRegular And InventoryInfo(LastClickedInvX, LastClickedInvY).Value >= DirectionCount Then
                        '--Rotate letters off back
                        InventoryInfo(LastClickedInvX, LastClickedInvY).Value = 1
                    End If
                Else
                    If InventoryInfo(LastClickedInvX, LastClickedInvY).Value > 0 Then
                        InventoryInfo(LastClickedInvX, LastClickedInvY).Value -= 1
                    Else
                        InventoryInfo(LastClickedInvX, LastClickedInvY).TerrainType = TerrainNull
                    End If
                End If
            Else
                If LeftClick Then
                    InventoryInfo(LastClickedInvX, LastClickedInvY).Value = 0
                    InventoryInfo(LastClickedInvX, LastClickedInvY).TerrainType = LastClickedLegend
                End If
            End If
        Else

        End If

        UpdateInventory()
    End Sub
#End Region

#Region " Support "
    Function Safe(ByVal x As Integer, ByVal y As Integer) As Boolean
        If x >= 0 And y >= 0 And x <= GridWidth And y <= GridHeight Then
            Return True
        Else
            Return False
        End If
    End Function

    Sub UpdateGrid()
        Dim i, j As Integer
        For j = 0 To GridHeight
            For i = 0 To GridWidth
                If BoxInfo(i, j).TerrainType >= 0 Then
                    TheBoxes(i, j).BackColor = ColorArray(BoxInfo(i, j).TerrainType)
                Else
                    '--Just in case, should not happen
                    TheBoxes(i, j).BackColor = ColorNull
                End If

                'TheBoxes(i, j).Appearance.FontData.SizeInPoints = BigFont
                If BoxInfo(i, j).Value > 0 Then
                    If BoxInfo(i, j).TerrainType = TerrainRegular Then
                        '--Directional Letter
                        TheBoxes(i, j).Text = DirectionArray(BoxInfo(i, j).Value)
                    Else
                        '--Value number if not zero
                        TheBoxes(i, j).Text = BoxInfo(i, j).Value
                    End If
                Else
                    TheBoxes(i, j).Text = ""
                End If
            Next
        Next
    End Sub

    Sub UpdateInventory()
        Dim i, j As Integer
        For j = 0 To InventoryHeight
            For i = 0 To InventoryWidth
                If InventoryInfo(i, j).TerrainType >= 0 Then
                    TheInventory(i, j).BackColor = ColorArray(InventoryInfo(i, j).TerrainType)
                Else
                    '--Unused inventory slot
                    TheInventory(i, j).BackColor = ColorNull
                End If

                'TheInventory(i, j).Appearance.FontData.SizeInPoints = BigFont
                If InventoryInfo(i, j).Value > 0 Then
                    If InventoryInfo(i, j).TerrainType = TerrainRegular Then
                        '--Directional Letter
                        TheInventory(i, j).Text = DirectionArray(InventoryInfo(i, j).Value)
                    Else
                        '--Value number if not zero
                        TheInventory(i, j).Text = InventoryInfo(i, j).Value
                    End If

                Else
                    TheInventory(i, j).Text = ""
                End If
            Next
        Next
    End Sub

    Function CountGoals() As Integer
        Dim i, j, goals As Integer
        goals = 0
        For j = 0 To GridHeight
            For i = 0 To GridWidth
                If BoxInfo(i, j).TerrainType = TerrainGoal Then
                    goals += 1
                End If
            Next
        Next
        Return goals
    End Function

    Function GatherTerrain(ByVal theTerrain As Integer) As ArrayList
        Dim thePairs As New ArrayList
        Dim i, j As Integer
        For j = 0 To GridHeight
            For i = 0 To GridWidth
                If BoxInfo(i, j).TerrainType = theTerrain Then
                    thePairs.Add(New Pair(i, j))
                End If
            Next
        Next
        Return thePairs
    End Function

    Sub ClearGoals()
        Dim i, j As Integer
        For j = 0 To GridHeight
            For i = 0 To GridWidth
                If BoxInfo(i, j).TerrainType = TerrainGoal Then
                    BoxInfo(i, j).Value = 0
                End If
            Next
        Next
        UpdateGrid()
    End Sub

    Sub ClearVisited()
        Dim i, j As Integer
        For j = 0 To GridHeight
            For i = 0 To GridWidth
                BoxInfo(i, j).MinDistance = MaxDistance
                BoxInfo(i, j).VisitedKey = -1
                BoxInfo(i, j).Parent.X = -1
                BoxInfo(i, j).Parent.Y = -1
            Next
        Next
    End Sub

    Sub ClearAll()
        Dim i, j As Integer
        For j = 0 To GridHeight
            For i = 0 To GridWidth
                BoxInfo(i, j).TerrainType = TerrainRegular
                BoxInfo(i, j).Value = 0
                BoxInfo(i, j).MinDistance = MaxDistance
                BoxInfo(i, j).VisitedKey = -1
                BoxInfo(i, j).Parent.X = -1
                BoxInfo(i, j).Parent.Y = -1
            Next
        Next
        UpdateGrid()
    End Sub

    Function GoalValidate(ByRef theGoalArray() As Pair) As Boolean
        Dim i, j, goals As Integer
        goals = 0
        For j = 0 To GridHeight
            For i = 0 To GridWidth
                If BoxInfo(i, j).TerrainType = TerrainGoal Then
                    goals += 1
                    If BoxInfo(i, j).Value < 1 Or BoxInfo(i, j).Value > theGoalArray.Length() Then
                        '--Too low or too high a number
                        Return False
                    Else
                        theGoalArray(BoxInfo(i, j).Value - 1) = New Pair(i, j)
                    End If
                End If
            Next
        Next
        If goals > 1 Then
            For i = 0 To theGoalArray.Length() - 1
                If IsDBNull(theGoalArray(i)) Then
                    '--Number Skipped
                    Return False
                ElseIf theGoalArray(i) Is Nothing Then
                    '--Number Skipped
                    Return False
                End If
            Next
            Return True
        Else
            '--Less than two goals given
            Return False
        End If
    End Function

    Private Sub PauseProgram(ByVal TimePeriod As Integer)
        Dim timeOut As DateTime = Now.AddMilliseconds(TimePeriod)
        Do
            'Keep the app from freezing and allow Windows to continue processing the applications messages.
            Application.DoEvents()
        Loop Until Now > timeOut
    End Sub
#End Region

#Region " Solve TSP (absolute shortest) "

    Sub SolveTSP()
        '--Create GoalList
        TotalGoals = CountGoals()
        If TotalGoals > 8 Then
            MsgBox("Too many goals to solve.", MsgBoxStyle.OKOnly, "Error")
            Return
        End If
        Dim LocalGoalList(TotalGoals - 1) As Pair
        Dim i, j, k As Integer
        k = 0
        For j = 0 To GridHeight
            For i = 0 To GridWidth
                If BoxInfo(i, j).TerrainType = TerrainGoal Then
                    LocalGoalList(k) = New Pair(i, j)
                    k += 1
                End If
            Next
        Next
        GoalList = LocalGoalList

        '--Fill Permutations array
        ListPermutations()

        '--Create Lookup Table
        Dim LookupTable(GridWidth, GridHeight) As Double
        For j = 0 To TotalGoals - 1
            For i = 0 To TotalGoals - 1
                If i <> j Then
                    TotalPathLength = 0.0
                    PathStack.Clear()
                    If Connect2Goals(i, j) Then
                        LookupTable(i, j) = TotalPathLength
                    Else
                        '-- Unconnectable
                        LookupTable(i, j) = MaxDistance
                    End If

                Else
                    LookupTable(i, j) = 0
                End If
            Next
        Next

        '--Check all paths
        Dim OverallMinDist As Double = MaxDistance
        Dim CurrentTotalDist As Double = 0
        Dim SolutionArray(TotalGoals - 1) As Integer
        Dim Beginning, Ending As Integer
        For i = 0 To Permutations.Count - 1
            CurrentTotalDist = 0.0
            For j = 0 To Permutations(i).Length - 2
                Beginning = Permutations(i)(j)
                Ending = Permutations(i)(j + 1)
                CurrentTotalDist += LookupTable(Beginning, Ending)
            Next
            If CurrentTotalDist < OverallMinDist Then
                OverallMinDist = CurrentTotalDist
                SolutionArray = Permutations(i)
            End If
        Next
        If OverallMinDist >= MaxDistance Then
            MsgBox("No path found.", MsgBoxStyle.OKOnly, "Error")
            Return
        End If

        '--Set Goallist to reflect best path
        Dim TempGoalList(TotalGoals - 1) As Pair
        Dim numberValue As Integer = 0
        For i = 0 To TotalGoals - 1
            TempGoalList(i) = New Pair(GoalList(SolutionArray(i)).X, GoalList(SolutionArray(i)).Y)
        Next
        GoalList = TempGoalList

        '--Run on solution to find path
        ConnectAllGoals()
    End Sub

    Sub ListPermutations()
        TotalGoals = CountGoals()
        Dim NumberArray(TotalGoals - 1) As Integer
        Dim i As Integer
        For i = 0 To TotalGoals - 1
            NumberArray(i) = i
        Next
        Permutations.Clear()
        HeapPermute(TotalGoals, NumberArray)

        '--Use for debugging
        'Dim readout As String = ""
        'For i = 0 To Permutations.Count - 1
        '    For j = 0 To Permutations(i).length - 1
        '        readout += Permutations(i)(j).ToString() + " "
        '    Next
        '    readout += ControlChars.NewLine
        'Next
        'MsgBox(readout, MsgBoxStyle.OKOnly, "Permutations")
    End Sub

    Sub HeapPermute(ByVal n As Integer, ByRef IntegerArray() As Integer)
        '--A. Levitin, Introduction to The Design & Analysis of Algorithms, Addison Wesley, 2003 pg 179
        If (n = 1) Then
            Dim FreshArray(TotalGoals - 1) As Integer
            Dim i As Integer
            For i = 0 To TotalGoals - 1
                FreshArray(i) = IntegerArray(i)
            Next
            Permutations.Add(FreshArray)
        Else
            Dim i As Integer
            For i = 0 To n - 1
                HeapPermute(n - 1, IntegerArray)
                If (n Mod 2 = 1) Then
                    '--if n is odd
                    Swap(0, n - 1, IntegerArray)
                Else
                    '--if n is even
                    Swap(i, n - 1, IntegerArray)
                End If
            Next
        End If
    End Sub

    Sub Swap(ByVal first As Integer, ByVal second As Integer, ByRef IntegerArray() As Integer)
        Dim temp As Integer = IntegerArray(first)
        IntegerArray(first) = IntegerArray(second)
        IntegerArray(second) = temp
    End Sub

    Sub AddItem(ByVal theString As String)
        Permutations.Add(theString)
    End Sub


#End Region

#Region " Solve "

    Sub SolvePuzzle()
        TotalGoals = CountGoals()
        Dim LocalGoalList(TotalGoals - 1) As Pair
        If GoalValidate(LocalGoalList) Then
            GoalList = LocalGoalList
            ConnectAllGoals()
        Else
            MsgBox("There must be at least 2 goals and they must be numbered properly.", MsgBoxStyle.OKOnly, "Error")
            Return
        End If
    End Sub

    Sub ConnectAllGoals()
        Dim i As Integer
        Dim readout As String = ""
        PathStack.Clear()
        TotalPathLength = 0.0
        For i = TotalGoals - 2 To 0 Step -1
            If Not Connect2Goals(i, i + 1) Then
                MsgBox("No path found.", MsgBoxStyle.OKOnly, "Error")
                Return
            End If
            'readout += (i + 1).ToString + " = (" + GoalList(i).X.tostring() + "," + GoalList(i).Y.tostring() + ")" + ControlChars.NewLine
        Next

        '--Results
        txtCurrentSolve.Text = TotalPathLength
        If TotalPathLength <= numTarget.Value Then
            MsgBox("Congratulations! You beat this level.", MsgBoxStyle.OKOnly, "Puzzle Solved!")
        End If


        PathList.Clear()
        While PathStack.Count > 0 'Not IsDBNull(PathStack.Peek)
            PathList.Add(PathStack.Pop)
        End While

        Dim PathLength As Integer = PathList.Count
        i = 0
        While i < PathLength
            If i > 0 Then
                If PathList(i).X = PathList(i - 1).X And PathList(i).Y = PathList(i - 1).Y Then
                    PathList.RemoveAt(i)
                    PathLength -= 1
                End If
            End If
            i += 1
            '    readout += (i + 1).ToString + " = (" + PathList(i).X.ToString() + "," + PathList(i).Y.ToString() + ") "
        End While
        'MsgBox(readout, MsgBoxStyle.OKOnly, "Goals")
    End Sub

    Function Connect2Goals(ByVal StartGoal As Integer, ByVal EndGoal As Integer) As Boolean
        ClearVisited()
        '--Recurse
        PathRecursion(GoalList(StartGoal).X, GoalList(StartGoal).Y, New Pair(-2, -2), 0.0)
        '--Handle Sinks
        Dim theSinks As New ArrayList
        theSinks = GatherTerrain(TerrainSink)
        Dim i As Integer
        For i = 0 To theSinks.Count - 1
            Dim theValue As Double = BoxInfo(theSinks(i).X, theSinks(i).Y).Value
            If theValue > 1 Then
                PathRecursion(theSinks(i).X, theSinks(i).Y, New Pair(GoalList(StartGoal).X, GoalList(StartGoal).Y), (theValue * StepSpeed))
            Else
                PathRecursion(theSinks(i).X, theSinks(i).Y, New Pair(GoalList(StartGoal).X, GoalList(StartGoal).Y), StepSpeed)
            End If
        Next

        '--Trace up the parents
        If BoxInfo(GoalList(EndGoal).X, GoalList(EndGoal).Y).VisitedKey > 0 Then
            Dim CurrentPair As New Pair(GoalList(EndGoal).X, GoalList(EndGoal).Y)
            PathStack.Push(CurrentPair)
            While (CurrentPair.X <> GoalList(StartGoal).X Or CurrentPair.Y <> GoalList(StartGoal).Y)
                '--Fill in pathstack by traveling up through parents
                CurrentPair = BoxInfo(CurrentPair.X, CurrentPair.Y).Parent
                PathStack.Push(New Pair(CurrentPair.X, CurrentPair.Y))
            End While
            TotalPathLength += BoxInfo(GoalList(EndGoal).X, GoalList(EndGoal).Y).MinDistance
        Else
            Return False
        End If
        Return True
    End Function

    Sub PathRecursion(ByVal X As Integer, ByVal Y As Integer, ByVal Parent As Pair, ByVal Distance As Double)
        If Safe(X, Y) Then
            If BoxInfo(X, Y).MinDistance > Distance Then
                If BoxInfo(X, Y).TerrainType = TerrainWall Or BoxInfo(X, Y).TerrainType = TerrainUnbreakable Then
                    If BoxInfo(X, Y).Value < 1 Then
                        '-- Hard wall
                        Return
                    Else
                        '-- Soft wall (pay penalty)
                        Distance += BoxInfo(X, Y).Value
                    End If
                End If

                '--Update internal
                BoxInfo(X, Y).Parent = Parent
                BoxInfo(X, Y).VisitedKey = 1
                BoxInfo(X, Y).MinDistance = Distance
                Dim LocalSS As Double = StepSpeed
                Dim LocalSSD As Double = StepSpeedDiagonal

                If BoxInfo(X, Y).TerrainType = TerrainFast Then
                    '--Fast
                    LocalSS = Math.Max(SafeDivide(LocalSS, BoxInfo(X, Y).Value), 0.0)
                    LocalSSD = Math.Max(SafeDivide(LocalSSD, BoxInfo(X, Y).Value), 0.0)
                End If
                If BoxInfo(X, Y).TerrainType = TerrainJump Then
                    '--Jump
                    PathRecursion(X + BoxInfo(X, Y).Value, Y, New Pair(X, Y), Distance + StepSpeed)
                    PathRecursion(X - BoxInfo(X, Y).Value, Y, New Pair(X, Y), Distance + StepSpeed)
                    PathRecursion(X, Y + BoxInfo(X, Y).Value, New Pair(X, Y), Distance + StepSpeed)
                    PathRecursion(X, Y - BoxInfo(X, Y).Value, New Pair(X, Y), Distance + StepSpeed)
                ElseIf BoxInfo(X, Y).TerrainType = TerrainTeleport Then
                    '--Teleport
                    Dim theTeleports As New ArrayList
                    theTeleports = GatherTerrain(TerrainTeleport)
                    Dim i As Integer
                    For i = 0 To theTeleports.Count - 1
                        If BoxInfo(X, Y).Value < 1 Then
                            PathRecursion(theTeleports(i).X, theTeleports(i).Y, New Pair(X, Y), Distance + StepSpeed)
                        ElseIf BoxInfo(X, Y).Value = BoxInfo(theTeleports(i).X, theTeleports(i).Y).Value Then
                            PathRecursion(theTeleports(i).X, theTeleports(i).Y, New Pair(X, Y), Distance + StepSpeed)
                        End If
                    Next
                End If

                If BoxInfo(X, Y).TerrainType = TerrainRegular And BoxInfo(X, Y).Value > 0 Then
                    'Limited Direction Recurse
                    Select Case (BoxInfo(X, Y).Value)
                        Case DirectionN
                            PathRecursion(X, Y - 1, New Pair(X, Y), Distance + LocalSS)
                        Case DirectionE
                            PathRecursion(X + 1, Y, New Pair(X, Y), Distance + LocalSS)
                        Case DirectionS
                            PathRecursion(X, Y + 1, New Pair(X, Y), Distance + LocalSS)
                        Case DirectionW
                            PathRecursion(X - 1, Y, New Pair(X, Y), Distance + LocalSS)
                        Case DirectionV
                            PathRecursion(X, Y + 1, New Pair(X, Y), Distance + LocalSS)
                            PathRecursion(X, Y - 1, New Pair(X, Y), Distance + LocalSS)
                        Case DirectionH
                            PathRecursion(X + 1, Y, New Pair(X, Y), Distance + LocalSS)
                            PathRecursion(X - 1, Y, New Pair(X, Y), Distance + LocalSS)
                        Case DirectionA
                            PathRecursion(X + 1, Y, New Pair(X, Y), Distance + LocalSS)
                            PathRecursion(X - 1, Y, New Pair(X, Y), Distance + LocalSS)
                            PathRecursion(X, Y + 1, New Pair(X, Y), Distance + LocalSS)
                            PathRecursion(X, Y - 1, New Pair(X, Y), Distance + LocalSS)
                    End Select
                Else
                    'Regular Recurse
                    PathRecursion(X + 1, Y, New Pair(X, Y), Distance + LocalSS)
                    PathRecursion(X - 1, Y, New Pair(X, Y), Distance + LocalSS)
                    PathRecursion(X, Y + 1, New Pair(X, Y), Distance + LocalSS)
                    PathRecursion(X, Y - 1, New Pair(X, Y), Distance + LocalSS)
                    PathRecursion(X + 1, Y + 1, New Pair(X, Y), Distance + LocalSSD)
                    PathRecursion(X + 1, Y - 1, New Pair(X, Y), Distance + LocalSSD)
                    PathRecursion(X - 1, Y + 1, New Pair(X, Y), Distance + LocalSSD)
                    PathRecursion(X - 1, Y - 1, New Pair(X, Y), Distance + LocalSSD)
                End If

            End If
        End If
    End Sub

    Sub ViewPath()
        Dim i As Integer
        For i = 0 To PathList.Count
            If i > 0 Then
                TheBoxes(PathList(i - 1).X, PathList(i - 1).y).BackColor = ColorArray(BoxInfo(PathList(i - 1).X, PathList(i - 1).y).TerrainType)
            End If
            If i < PathList.Count Then
                TheBoxes(PathList(i).X, PathList(i).y).BackColor = ColorSelected
                PauseProgram(200)
            End If
        Next
    End Sub

#End Region

#Region " Buttons "
    Private Sub btSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btSave.Click
        SaveLevel()
    End Sub

    Private Sub btLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btLoad.Click
        LoadLevel()
    End Sub

    Private Sub btNumberGoals_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btNumberGoals.Click
        TotalGoals = CountGoals()
        If TotalGoals > 1 Then
            rbPlay.Checked = True
            ClearGoals()
            NumberingGoals = True
            GoalNumber = 0
        Else
            MsgBox("You must have at least two goals (one to start on and one to finish on).", MsgBoxStyle.OKOnly, "Error")
        End If
    End Sub

    Private Sub Solve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btSolve.Click
        SolvePuzzle()
    End Sub

    Private Sub btSolveTSP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btSolveTSP.Click
        SolveTSP()
    End Sub

    Private Sub btViewPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btViewPath.Click
        If PathList.Count > 0 Then
            ViewPath()
        End If
    End Sub

    Private Sub btClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClear.Click
        ClearAll()
    End Sub

    Private Sub rbPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPlay.Click
        UpdateRadioClick()
    End Sub

    Private Sub rbPlay_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPlay.CheckedChanged
        UpdateRadioClick()
    End Sub

    Private Sub rbEdit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEdit.CheckedChanged
        UpdateRadioClick()
    End Sub

    Private Sub btReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btReset.Click
        GoalNumber = 0
        LoadBackup()
    End Sub

    Private Sub rbEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEdit.Click
        UpdateRadioClick()
    End Sub

    Sub UpdateRadioClick()
        If rbPlay.Checked Then
            '--Play Mode
            txtLevelName.ReadOnly = True
            numTarget.ReadOnly = True
            numTarget.Increment = 0
            cbPresetOrder.Enabled = False
            btSave.Enabled = False
            btClear.Enabled = False
        Else
            '--Edit Mode
            txtLevelName.ReadOnly = False
            numTarget.ReadOnly = False
            numTarget.Increment = 1
            cbPresetOrder.Enabled = True
            btSave.Enabled = True
            btClear.Enabled = True
        End If
    End Sub

    Private Sub cbPresetOrder_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPresetOrder.CheckedChanged
        UpdatePreset()
    End Sub

    Sub UpdatePreset()
        If cbPresetOrder.Checked Then
            btNumberGoals.Enabled = False
            btSolveTSP.Enabled = False
        Else
            btNumberGoals.Enabled = True
            btSolveTSP.Enabled = True
        End If
    End Sub

    Private Sub btRules_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btRules.Click
        MsgBox("See the Instruction Manual and the Delivergrid Report in the main file folder for more information.", MsgBoxStyle.OKOnly, "Rules")
    End Sub

#End Region

#Region " Save and Load "
    Private Sub SaveLevel()
        Dim TheSaveDialog As New SaveFileDialog
        Dim Parser As New Parser
        Dim i, j As Integer
        TheSaveDialog.InitialDirectory = Parser.AddSlash(Application.StartupPath) & "MapSaves"
        TheSaveDialog.Filter = "Delivergrid Maps (.map)|*.map"
        TheSaveDialog.DefaultExt = ".map"
        If TheSaveDialog.ShowDialog = DialogResult.OK Then
            Dim fi As New FileInfo(TheSaveDialog.FileName)
            If fi.Exists Then
                fi.Delete()
            End If
            '-- Open File
            Dim sw As StreamWriter = New StreamWriter(TheSaveDialog.FileName)
            '-- Record Preset Order
            sw.Write(cbPresetOrder.Checked.ToString & ControlChars.Lf)
            '-- Record Target Distance
            sw.Write(numTarget.Value.ToString & ControlChars.Lf)
            '-- Record Each Square
            For j = 0 To InventoryHeight
                For i = 0 To InventoryWidth
                    sw.Write(InventoryInfo(i, j).BuildSaveString & ControlChars.Lf)
                Next
            Next
            For j = 0 To GridHeight
                For i = 0 To GridWidth
                    sw.Write(BoxInfo(i, j).BuildSaveString & ControlChars.Lf)
                Next
            Next
            '-- Close File
            sw.Close()
            txtLevelName.Text = Parser.GetLevelName(TheSaveDialog.FileName)
            '--Save Level info as backup
            SaveBackup()
            MsgBox("Map saved successfully.", , "Map Save")
        End If

    End Sub
    Private Sub LoadLevel()
        '--Cleanup
        rbPlay.Checked = True
        LastClickedInvX = InventoryWidth
        LastClickedInvY = InventoryHeight
        PathList.Clear()

        Dim TheLoadDialog As New OpenFileDialog
        Dim Parser As New Parser
        TheLoadDialog.InitialDirectory = Parser.AddSlash(Application.StartupPath) & "MapSaves"
        TheLoadDialog.Filter = "Delivergrid Maps (.map)|*.map"
        TheLoadDialog.DefaultExt = ".map"
        If TheLoadDialog.ShowDialog = DialogResult.OK Then
            Dim i As Integer = 0
            Dim j As Integer = 0
            Try
                ' Create an instance of StreamReader to read from a file.
                Dim sr As StreamReader = New StreamReader(TheLoadDialog.FileName)
                '-- Record first line as Preset Order
                cbPresetOrder.Checked = sr.ReadLine
                ' Read in second line as Target Distance
                numTarget.Value = sr.ReadLine

                ' Read and display the lines from the file until the end 
                ' of the file is reached.
                For j = 0 To InventoryHeight
                    For i = 0 To InventoryWidth
                        InventoryInfo(i, j).ReadLoadString(sr.ReadLine)
                    Next
                Next
                For j = 0 To GridHeight
                    For i = 0 To GridWidth
                        BoxInfo(i, j).ReadLoadString(sr.ReadLine)
                    Next
                Next
                sr.Close()
                txtLevelName.Text = Parser.GetLevelName(TheLoadDialog.FileName)
                '--Save Level info as backup
                SaveBackup()
            Catch Ex As Exception
                MsgBox("Couldn't read file.", , "Load Error")
            End Try
        End If

    End Sub

    Sub SaveBackup()
        Dim TempArray(GridWidth, GridHeight) As TerrainSquare
        'Array.Copy(BoxInfo, CurrentLevel.BackupGrid, BoxInfo.Length)
        'Array.Copy(InventoryInfo, CurrentLevel.BackupInventory, InventoryInfo.Length)
        'CurrentLevel.BackupGrid = BoxInfo
        'CurrentLevel.BackupInventory = InventoryInfo
        Dim i, j As Integer
        For j = 0 To GridHeight
            For i = 0 To GridWidth
                CurrentLevel.BackupGrid(i, j) = New TerrainSquare(i, j, BoxInfo(i, j).TerrainType, BoxInfo(i, j).Value)
            Next
        Next
        For j = 0 To InventoryHeight
            For i = 0 To InventoryWidth
                CurrentLevel.BackupInventory(i, j) = New TerrainSquare(i, j, InventoryInfo(i, j).TerrainType, InventoryInfo(i, j).Value)
            Next
        Next
        CurrentLevel.TargetDistance = numTarget.Value
        UpdateGrid()
        UpdateInventory()
    End Sub

    Sub LoadBackup()
        'Array.Copy(CurrentLevel.BackupGrid, BoxInfo, CurrentLevel.BackupGrid.Length)
        'Array.Copy(CurrentLevel.BackupInventory, InventoryInfo, CurrentLevel.BackupInventory.Length)
        'BoxInfo = CurrentLevel.BackupGrid
        'InventoryInfo = CurrentLevel.BackupInventory
        Dim i, j As Integer
        For j = 0 To GridHeight
            For i = 0 To GridWidth
                BoxInfo(i, j) = New TerrainSquare(i, j, CurrentLevel.BackupGrid(i, j).TerrainType, CurrentLevel.BackupGrid(i, j).Value)
            Next
        Next
        For j = 0 To InventoryHeight
            For i = 0 To InventoryWidth
                InventoryInfo(i, j) = New TerrainSquare(i, j, CurrentLevel.BackupInventory(i, j).TerrainType, CurrentLevel.BackupInventory(i, j).Value)
            Next
        Next
        numTarget.Value = CurrentLevel.TargetDistance
        UpdateGrid()
        UpdateInventory()
    End Sub
#End Region


End Class
