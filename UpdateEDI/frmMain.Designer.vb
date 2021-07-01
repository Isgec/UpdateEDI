<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    Try
      If disposing AndAlso components IsNot Nothing Then
        components.Dispose()
      End If
    Finally
      MyBase.Dispose(disposing)
    End Try
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.split1 = New System.Windows.Forms.SplitContainer()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.cmdStop = New System.Windows.Forms.Button()
    Me.cmdStart = New System.Windows.Forms.Button()
    Me.split2 = New System.Windows.Forms.SplitContainer()
    Me.ListBox1 = New System.Windows.Forms.ListBox()
    Me.ListBox2 = New System.Windows.Forms.ListBox()
    CType(Me.split1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.split1.Panel1.SuspendLayout()
    Me.split1.Panel2.SuspendLayout()
    Me.split1.SuspendLayout()
    CType(Me.split2, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.split2.Panel1.SuspendLayout()
    Me.split2.Panel2.SuspendLayout()
    Me.split2.SuspendLayout()
    Me.SuspendLayout()
    '
    'split1
    '
    Me.split1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.split1.Location = New System.Drawing.Point(0, 0)
    Me.split1.Name = "split1"
    Me.split1.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'split1.Panel1
    '
    Me.split1.Panel1.Controls.Add(Me.Label1)
    Me.split1.Panel1.Controls.Add(Me.cmdStop)
    Me.split1.Panel1.Controls.Add(Me.cmdStart)
    '
    'split1.Panel2
    '
    Me.split1.Panel2.Controls.Add(Me.split2)
    Me.split1.Size = New System.Drawing.Size(743, 322)
    Me.split1.SplitterDistance = 40
    Me.split1.TabIndex = 3
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(3, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(50, 13)
    Me.Label1.TabIndex = 4
    Me.Label1.Text = "Message"
    '
    'cmdStop
    '
    Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdStop.Enabled = False
    Me.cmdStop.Location = New System.Drawing.Point(656, 4)
    Me.cmdStop.Name = "cmdStop"
    Me.cmdStop.Size = New System.Drawing.Size(75, 23)
    Me.cmdStop.TabIndex = 3
    Me.cmdStop.Text = "Stop"
    Me.cmdStop.UseVisualStyleBackColor = True
    '
    'cmdStart
    '
    Me.cmdStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdStart.Location = New System.Drawing.Point(555, 4)
    Me.cmdStart.Name = "cmdStart"
    Me.cmdStart.Size = New System.Drawing.Size(75, 23)
    Me.cmdStart.TabIndex = 2
    Me.cmdStart.Text = "Start"
    Me.cmdStart.UseVisualStyleBackColor = True
    '
    'split2
    '
    Me.split2.Dock = System.Windows.Forms.DockStyle.Fill
    Me.split2.Location = New System.Drawing.Point(0, 0)
    Me.split2.Name = "split2"
    '
    'split2.Panel1
    '
    Me.split2.Panel1.Controls.Add(Me.ListBox1)
    '
    'split2.Panel2
    '
    Me.split2.Panel2.Controls.Add(Me.ListBox2)
    Me.split2.Size = New System.Drawing.Size(743, 278)
    Me.split2.SplitterDistance = 222
    Me.split2.TabIndex = 0
    '
    'ListBox1
    '
    Me.ListBox1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.ListBox1.FormattingEnabled = True
    Me.ListBox1.Location = New System.Drawing.Point(0, 0)
    Me.ListBox1.Name = "ListBox1"
    Me.ListBox1.Size = New System.Drawing.Size(222, 278)
    Me.ListBox1.TabIndex = 0
    '
    'ListBox2
    '
    Me.ListBox2.Dock = System.Windows.Forms.DockStyle.Fill
    Me.ListBox2.FormattingEnabled = True
    Me.ListBox2.Location = New System.Drawing.Point(0, 0)
    Me.ListBox2.Name = "ListBox2"
    Me.ListBox2.Size = New System.Drawing.Size(517, 278)
    Me.ListBox2.TabIndex = 1
    '
    'frmMain
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(743, 322)
    Me.Controls.Add(Me.split1)
    Me.Name = "frmMain"
    Me.Text = "Update EDI"
    Me.split1.Panel1.ResumeLayout(False)
    Me.split1.Panel1.PerformLayout()
    Me.split1.Panel2.ResumeLayout(False)
    CType(Me.split1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.split1.ResumeLayout(False)
    Me.split2.Panel1.ResumeLayout(False)
    Me.split2.Panel2.ResumeLayout(False)
    CType(Me.split2, System.ComponentModel.ISupportInitialize).EndInit()
    Me.split2.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents split1 As SplitContainer
  Friend WithEvents Label1 As Label
  Friend WithEvents cmdStop As Button
  Friend WithEvents cmdStart As Button
  Friend WithEvents split2 As SplitContainer
  Friend WithEvents ListBox1 As ListBox
  Friend WithEvents ListBox2 As ListBox
End Class
