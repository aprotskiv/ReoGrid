/*****************************************************************************
 * 
 * ReoGrid - .NET Spreadsheet Control
 * 
 * https://reogrid.net/
 *
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
 * PURPOSE.
 *
 * Author: Jingwood <jingwood at unvell.com>
 *
 * Copyright (c) 2012-2025 Jingwood <jingwood at unvell.com>
 * Copyright (c) 2012-2025 UNVELL Inc. All rights reserved.
 * 
 ****************************************************************************/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;

using unvell.Common;

using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Views;
using unvell.ReoGrid.Interaction;

#if WINFORM
using RGFloat = System.Single;
using RGPointF = System.Drawing.PointF;
using IntOrDouble = System.Int32;
//using ReoGridControl = unvell.ReoGrid.ReoGridControl;

#elif WPF
using RGFloat = System.Double;
using RGPoint = System.Windows.Point;
using RGPointF = System.Windows.Point;
using IntOrDouble = System.Double;
//using ReoGridControl = unvell.ReoGrid.ReoGridControl;

#elif ANDROID
using RGFloat = System.Single;
using RGPoint = Android.Graphics.Point;
using IntOrDouble = System.Int32;
using ReoGridControl = unvell.ReoGrid.ReoGridView;

#elif iOS
using RGFloat = System.Double;
using RGPointF = CoreGraphics.CGPoint;
using IntOrDouble = System.Double;
using ReoGridControl = unvell.ReoGrid.ReoGridView;

#endif // WPF

#if WINFORM
using Cursor = System.Windows.Forms.Cursor;
//using Cursors = System.Windows.Forms.Cursor;
#elif WPF
using Cursor = System.Windows.Input.Cursor;
//using Cursors = System.Windows.Input.Cursors;
#endif // WPF

using unvell.ReoGrid.Main;
using unvell.ReoGrid.Rendering;
using System.Windows.Input;

namespace unvell.ReoGrid
{

#if WINFORM || WPF
  partial class ReoGridControl
#elif ANDROID || iOS
	partial class ReoGridView
#endif // ANDROID || iOS
  {
    private IRenderer renderer;

    #region Initialize

    static ReoGridControl()
    {
#if NETCOREAPP3_1_OR_GREATER
				Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif // NETCOREAPP3_1_OR_GREATER
    }

    private void InitControl()
    {
#if WINFORM || WPF
      // initialize cursors
      // normal grid selector
      this.builtInCellsSelectionCursor = LoadCursorFromResource(unvell.ReoGrid.Properties.Resources.grid_select);
      this.internalCurrentCursor = builtInCellsSelectionCursor;

      // cell picking
      this.defaultPickRangeCursor = LoadCursorFromResource(unvell.ReoGrid.Properties.Resources.pick_range);

      // full-row and full-col selector
      this.builtInFullColSelectCursor = LoadCursorFromResource(unvell.ReoGrid.Properties.Resources.full_col_select);
      this.builtInFullRowSelectCursor = LoadCursorFromResource(unvell.ReoGrid.Properties.Resources.full_row_select);

      this.builtInEntireSheetSelectCursor = this.builtInCellsSelectionCursor;

      this.builtInCrossCursor = LoadCursorFromResource(unvell.ReoGrid.Properties.Resources.cross);
#endif // WINFORM || WPF

      this.ControlStyle = ControlAppearanceStyle.CreateDefaultControlStyle();
    }

    private void InitWorkbook(IControlAdapter adapter)
    {
      // create workbook
      this.workbook = new Workbook(adapter);

      #region Workbook Event Attach
      this.workbook.WorksheetCreated += (s, e) =>
      {
        this.WorksheetCreated?.Invoke(this, e);
      };

      this.workbook.WorksheetInserted += (s, e) =>
      {
        this.WorksheetInserted?.Invoke(this, e);
      };

      this.workbook.WorksheetRemoved += (s, e) =>
      {
        this.ClearActionHistoryForWorksheet(e.Worksheet);

        if (this.workbook.worksheets.Count > 0)
        {
          int index = this.sheetTab.SelectedIndex;

          if (index >= this.workbook.worksheets.Count)
          {
            index = this.workbook.worksheets.Count - 1;
          }

          this.sheetTab.SelectedIndex = index;
          this.currentWorksheet = this.workbook.worksheets[this.sheetTab.SelectedIndex];
        }
        else
        {
          this.sheetTab.SelectedIndex = -1;
          this.currentWorksheet = null;
        }

        this.WorksheetRemoved?.Invoke(this, e);
      };

      this.workbook.WorksheetNameChanged += (s, e) =>
      {
        this.WorksheetNameChanged?.Invoke(this, e);
      };

      this.workbook.SettingsChanged += (s, e) =>
      {
        if (this.workbook.HasSettings(WorkbookSettings.View_ShowSheetTabControl))
        {
          ShowSheetTabControl();
        }
        else
        {
          HideSheetTabControl();
        }

        if (this.workbook.HasSettings(WorkbookSettings.View_ShowHorScroll))
        {
          ShowHorScrollBar();
        }
        else
        {
          HideHorScrollBar();
        }

        if (this.workbook.HasSettings(WorkbookSettings.View_ShowVerScroll))
        {
          ShowVerScrollBar();
        }
        else
        {
          HideVerScrollBar();
        }

        this.SettingsChanged?.Invoke(this, null);
      };

      this.workbook.ExceptionHappened += Workbook_ErrorHappened;

      #endregion // Workbook Event Attach

#if EX_SCRIPT
				this.workbook.SRMInitialized += (s, e) =>
						{
							if (this.workbook.workbookObj != null)
							{
									this.workbook.workbookObj.ControlInstance = this;
							}
						};
#endif // EX_SCRIPT

      // create and set default worksheet
      this.workbook.AddWorksheet(this.workbook.CreateWorksheet());

      //RefreshWorksheetTabs();
      this.CurrentWorksheet = this.workbook.worksheets[0];

      this.sheetTab.SelectedIndexChanged += (s, e) =>
      {
        if (this.sheetTab.SelectedIndex >= 0 && this.sheetTab.SelectedIndex < this.workbook.worksheets.Count)
        {
          this.CurrentWorksheet = this.workbook.worksheets[this.sheetTab.SelectedIndex];
        }
      };

      this.workbook.WorkbookLoaded += (s, e) =>
      {
        if (this.workbook.worksheets.Count <= 0)
        {
          this.currentWorksheet = null;
        }
        else
        {
          if (this.currentWorksheet != this.workbook.worksheets[0])
          {
            this.currentWorksheet = this.workbook.worksheets[0];
          }
          else
          {
            this.currentWorksheet.UpdateViewportControllBounds();
          }
        }

        this.WorkbookLoaded?.Invoke(s, e);
      };

      this.workbook.WorkbookSaved += (s, e) =>
      {
        this.WorkbookSaved?.Invoke(s, e);
      };


    }

    #endregion // Initialize

    #region Memory Workbook
    /// <summary>
    /// Create an instance of ReoGrid workbook in memory. <br/>
    /// The memory workbook is the non-GUI version of ReoGrid control, which can do almost all operations, 
    /// such as reading and saving from Excel file, RGF file, changing data, formulas, styles, borders and etc.
    /// </summary>
    /// <returns>Instance of memory workbook.</returns>
    public static IWorkbook CreateMemoryWorkbook()
    {
      var workbook = new Workbook(null);

      var defaultWorksheet = workbook.CreateWorksheet();
      workbook.AddWorksheet(defaultWorksheet);

      return workbook;
    }
    #endregion // Memory Workbook

    #region Workbook & Worksheet

    public IDataActionFactory DataActionFactory { get; set; } = new DefaultDataActionFactory();

    public IPartialGridFactory PartialGridFactory { get; set; } = new DefaultPartialGridFactory();

    private Workbook workbook;

    #region Save & Load
    /// <summary>
    /// Save workbook into file
    /// </summary>
    /// <param name="path">Full file path to save workbook</param>
    /// <param name="fileFormat">Specified file format used to save workbook</param>
    public void Save(string path)
    {
      this.Save(path, IO.FileFormat._Auto);
    }

    /// <summary>
    /// Save workbook into file
    /// </summary>
    /// <param name="path">Full file path to save workbook</param>
    /// <param name="fileFormat">Specified file format used to save workbook</param>
    public void Save(string path, IO.FileFormat fileFormat)
    {
      this.Save(path, fileFormat, Encoding.Default);
    }

    /// <summary>
    /// Save workbook into file
    /// </summary>
    /// <param name="path">Full file path to save workbook</param>
    /// <param name="fileFormat">Specified file format used to save workbook</param>
    /// <param name="encoding">Encoding used to read plain-text from resource. (Optional)</param>
    public void Save(string path, IO.FileFormat fileFormat, Encoding encoding)
    {
      this.workbook.Save(path, fileFormat, encoding);
    }

    /// <summary>
    /// Save workbook into stream with specified format
    /// </summary>
    /// <param name="stream">Stream to output data of workbook</param>
    /// <param name="fileFormat">Specified file format used to save workbook</param>
    public void Save(Stream stream, unvell.ReoGrid.IO.FileFormat fileFormat)
    {
      this.workbook.Save(stream, fileFormat, Encoding.Default);
    }

    /// <summary>
    /// Save workbook into stream with specified format
    /// </summary>
    /// <param name="stream">Stream to output data of workbook</param>
    /// <param name="fileFormat">Specified file format used to save workbook</param>
    /// <param name="encoding">Encoding used to read plain-text from resource. (Optional)</param>
    public void Save(Stream stream, unvell.ReoGrid.IO.FileFormat fileFormat, Encoding encoding)
    {
      this.workbook.Save(stream, fileFormat, encoding);
    }

    /// <summary>
    /// Load workbook from file by specified path.
    /// </summary>
    /// <param name="path">Path to open file and read data.</param>
    public void Load(string path)
    {
      this.Load(path, IO.FileFormat._Auto, Encoding.Default);
    }

    /// <summary>
    /// Load workbook from file by specified path.
    /// </summary>
    /// <param name="path">Path to open file and read data.</param>
    /// <param name="fileFormat">Flag used to determine what format should be used to read data from file.</param>
    public void Load(string path, IO.FileFormat fileFormat)
    {
      this.Load(path, fileFormat, Encoding.Default);
    }

    /// <summary>
    /// Load workbook from file with specified format
    /// </summary>
    /// <param name="path">Path to open file and read data.</param>
    /// <param name="fileFormat">Flag used to determine what format should be used to read data from file.</param>
    /// <param name="encoding">Encoding used to read plain-text from resource. (Optional)</param>
    public void Load(string path, IO.FileFormat fileFormat, Encoding encoding)
    {
      this.workbook.Load(path, fileFormat, encoding);
    }

    /// <summary>
    /// Load workbook from stream with specified format.
    /// </summary>
    /// <param name="stream">Stream to read data of workbook.</param>
    /// <param name="fileFormat">Flag used to determine what format should be used to read data from file.</param>
    public void Load(Stream stream, unvell.ReoGrid.IO.FileFormat fileFormat)
    {
      this.Load(stream, fileFormat, Encoding.Default);
    }

    /// <summary>
    /// Load workbook from stream with specified format.
    /// </summary>
    /// <param name="stream">Stream to read data of workbook.</param>
    /// <param name="fileFormat">Flag used to determine what format should be used to read data from file.</param>
    /// <param name="encoding">Encoding used to read plain-text data from specified stream.</param>
    public void Load(Stream stream, unvell.ReoGrid.IO.FileFormat fileFormat, Encoding encoding)
    {
      this.workbook.Load(stream, fileFormat, encoding);

      if (this.workbook.worksheets.Count > 0)
      {
        this.CurrentWorksheet = this.workbook.worksheets[0];
      }
    }
    #endregion // Save & Load

    /// <summary>
    /// Event raised when workbook loaded from stream or file.
    /// </summary>
    public event EventHandler WorkbookLoaded;

    /// <summary>
    /// Event raised when workbook saved into stream or file.
    /// </summary>
    public event EventHandler WorkbookSaved;

    #region Worksheet Management

    private Worksheet currentWorksheet;

    /// <summary>
    /// Get or set the current worksheet
    /// </summary>
    public Worksheet CurrentWorksheet
    {
      get
      {
        return this.currentWorksheet;
      }
      set
      {
        if (value == null) throw new ArgumentNullException("cannot set current worksheet to null");

        if (this.currentWorksheet != value)
        {
          if (this.currentWorksheet != null && this.currentWorksheet.IsEditing)
          {
            this.currentWorksheet.EndEdit(EndEditReason.NormalFinish);
          }

          this.currentWorksheet = value;

          // update bounds for viewport of worksheet
          this.currentWorksheet.UpdateViewportControllBounds();

          // update bounds for viewport of worksheet
          if (this.currentWorksheet.ViewportController is IScrollableViewportController scrollableViewportController)
          {
            scrollableViewportController.SynchronizeScrollBar();
          }

          this.CurrentWorksheetChanged?.Invoke(this, null);

          this.sheetTab.SelectedIndex = GetWorksheetIndex(this.currentWorksheet);
          this.sheetTab.ScrollToItem(this.sheetTab.SelectedIndex);

          this.adapter.Invalidate();
        }
      }
    }

    /// <summary>
    /// Create new instance of worksheet with default available name. (e.g. Sheet1, Sheet2 ...)
    /// </summary>
    /// <returns>Instance of worksheet to be created.</returns>
    /// <remarks>This method creates a new worksheet, but doesn't add it into the collection of worksheet.
    /// Worksheet will only be available until adding into a workbook, by using these methods:
    /// <code>InsertWorksheet</code>, <code>Worksheets.Add</code> or <code>Worksheets.Insert</code>
    /// </remarks>
    public Worksheet CreateWorksheet()
    {
      return this.CreateWorksheet(null);
    }

    /// <summary>
    /// Create new instance of worksheet.
    /// </summary>
    /// <param name="name">name of new worksheet to be created. 
    /// If name is null, ReoGrid will find an available name automatically. e.g. 'Sheet1', 'Sheet2'...</param>
    /// <returns>instance of worksheet to be created</returns>
    /// <remarks>This method creates a new worksheet, but doesn't add it into the collection of worksheet.
    /// Worksheet will only be available until adding into a workbook, by using these methods:
    /// <code>InsertWorksheet</code>, <code>Worksheets.Add</code> or <code>Worksheets.Insert</code>
    /// </remarks>
    public Worksheet CreateWorksheet(string name)
    {
      return this.workbook.CreateWorksheet(name);
    }

    /// <summary>
    /// Add specified worksheet into this workbook
    /// </summary>
    /// <param name="sheet">worksheet to be added</param>
    public void AddWorksheet(Worksheet sheet)
    {
      this.workbook.AddWorksheet(sheet);
    }

    /// <summary>
    /// Create and append a new instance of worksheet into workbook.
    /// </summary>
    /// <param name="name">Optional name for new worksheet.</param>
    /// <returns>Instance of created new worksheet.</returns>
    public Worksheet NewWorksheet(string name = null)
    {
      var worksheet = this.CreateWorksheet(name);

      this.AddWorksheet(worksheet);

      return worksheet;
    }

    /// <summary>
    /// Insert specified worksheet into this workbook.
    /// </summary>
    /// <param name="index">position of zero-based number of worksheet used to insert specified worksheet.</param>
    /// <param name="sheet">worksheet to be inserted.</param>
    public void InsertWorksheet(int index, Worksheet sheet)
    {
      this.workbook.InsertWorksheet(index, sheet);
    }

    /// <summary>
    /// Remove worksheet from this workbook by specified index.
    /// </summary>
    /// <param name="index">zero-based number of worksheet to be removed.</param>
    /// <returns>true if specified worksheet can be found and removed successfully.</returns>
    public bool RemoveWorksheet(int index)
    {
      return this.workbook.RemoveWorksheet(index);
    }

    /// <summary>
    /// Remove worksheet from this workbook.
    /// </summary>
    /// <param name="sheet">worksheet to be removed.</param>
    /// <returns>true if specified worksheet can be found and removed successfully.</returns>
    public bool RemoveWorksheet(Worksheet sheet)
    {
      return this.workbook.RemoveWorksheet(sheet);
    }

    /// <summary>
    /// Create a cloned worksheet and put into specified position.
    /// </summary>
    /// <param name="index">Index of source worksheet to be copied</param>
    /// <param name="newIndex">Target index used to insert the copied worksheet</param>
    /// <param name="newName">Name for new worksheet, set as null to use a default worksheet name e.g. Sheet1, Sheet2...</param>
    /// <returns>New instance of copid worksheet</returns>
    public Worksheet CopyWorksheet(int index, int newIndex, string newName = null)
    {
      return this.workbook.CopyWorksheet(index, newIndex, newName);
    }

    /// <summary>
    /// Create a cloned worksheet and put into specified position.
    /// </summary>
    /// <param name="sheet">Source worksheet to be copied, the worksheet must be already added into this workbook</param>
    /// <param name="newIndex">Target index used to insert the copied worksheet</param>
    /// <param name="newName">Name for new worksheet, set as null to use a default worksheet name e.g. Sheet1, Sheet2...</param>
    /// <returns>New instance of copid worksheet</returns>
    public Worksheet CopyWorksheet(Worksheet sheet, int newIndex, string newName = null)
    {
      return this.workbook.CopyWorksheet(sheet, newIndex, newName);
    }

    /// <summary>
    /// Move worksheet from a position to another position.
    /// </summary>
    /// <param name="index">Worksheet in this position to be moved</param>
    /// <param name="newIndex">Target position moved to</param>
    /// <returns>Instance of moved worksheet</returns>
    public Worksheet MoveWorksheet(int index, int newIndex)
    {
      return this.workbook.MoveWorksheet(index, newIndex);
    }

    /// <summary>
    /// Create a cloned worksheet and put into specified position.
    /// </summary>
    /// <param name="sheet">Instance of worksheet to be moved, the worksheet must be already added into this workbook.</param>
    /// <param name="newIndex">Zero-based target position moved to.</param>
    /// <returns>Instance of moved worksheet.</returns>
    public Worksheet MoveWorksheet(Worksheet sheet, int newIndex)
    {
      return this.workbook.MoveWorksheet(sheet, newIndex);
    }

    /// <summary>
    /// Get index of specified worksheet from the collection in this workbook
    /// </summary>
    /// <param name="sheet">Worksheet to get.</param>
    /// <returns>zero-based number of worksheet in this workbook's collection</returns>
    public int GetWorksheetIndex(Worksheet sheet)
    {
      return this.workbook.GetWorksheetIndex(sheet);
    }

    /// <summary>
    /// Get the index of specified worksheet by name from workbook.
    /// </summary>
    /// <param name="sheet">Name of worksheet to get.</param>
    /// <returns>Zero-based number of worksheet in worksheet collection of workbook. Returns -1 if not found.</returns>
    public int GetWorksheetIndex(string name)
    {
      return this.workbook.GetWorksheetIndex(name);
    }

    /// <summary>
    /// Find worksheet by specified name.
    /// </summary>
    /// <param name="name">Name to find worksheet.</param>
    /// <returns>Instance of worksheet that is found by specified name; otherwise return null.</returns>
    public Worksheet GetWorksheetByName(string name)
    {
      return this.workbook.GetWorksheetByName(name);
    }

    /// <summary>
    /// Get the collection of worksheet.
    /// </summary>
    //[System.ComponentModel.Editor(typeof(WinForm.Designer.WorkbookEditor),
    //	typeof(System.Drawing.Design.UITypeEditor))]
    public WorksheetCollection Worksheets
    {
      get { return this.workbook.Worksheets; }
    }

    /// <summary>
    /// Event raised when current worksheet is changed.
    /// </summary>
    public event EventHandler CurrentWorksheetChanged;

    /// <summary>
    /// Event raised when worksheet is created.
    /// </summary>
    public event EventHandler<WorksheetCreatedEventArgs> WorksheetCreated;

    /// <summary>
    /// Event raised when worksheet is inserted into this workbook.
    /// </summary>
    public event EventHandler<WorksheetInsertedEventArgs> WorksheetInserted;

    /// <summary>
    /// Event raised when worksheet is removed from this workbook.
    /// </summary>
    public event EventHandler<WorksheetRemovedEventArgs> WorksheetRemoved;

    /// <summary>
    /// Event raised when the name of worksheet managed by this workbook is changed.
    /// </summary>
    public event EventHandler<WorksheetNameChangingEventArgs> WorksheetNameChanged;

    /// <summary>
    /// Event raised when background color of worksheet name is changed.
    /// </summary>
    public event EventHandler<WorksheetEventArgs> WorksheetNameBackColorChanged;

    /// <summary>
    /// Event raised when text color of worksheet name is changed.
    /// </summary>
    public event EventHandler<WorksheetEventArgs> WorksheetNameTextColorChanged;

    #endregion // Worksheet Management

    /// <summary>
    /// Determine whether or not this workbook is read-only (Reserved v0.8.8)
    /// </summary>
    [Description("Determine whether or not this workbook is read-only")]
    [DefaultValue(false)]
    public bool Readonly
    {
      get
      {
        return this.workbook.Readonly;
      }
      set
      {
        this.workbook.Readonly = value;
      }
    }

    /// <summary>
    /// Reset control and workbook (remove all worksheets and put one new)
    /// </summary>
    public void Reset()
    {
      this.workbook.Reset();

      this.CurrentWorksheet = this.workbook.worksheets[0];
    }

    /// <summary>
    /// Check whether or not current workbook is empty (all worksheets don't have any cells)
    /// </summary>
    public bool IsWorkbookEmpty
    {
      get
      {
        return this.workbook.IsEmpty;
      }
    }

    #endregion // Workbook & Worksheet

    #region Actions

    private Lazy<IActionController> _actionControllerOverride;
    public Lazy<IActionController> ActionControllerOverride
    {
      get
      {
        if (_actionControllerOverride == null)
        {
          _actionControllerOverride = new Lazy<IActionController>(() => new ActionController(this));
        }
        return _actionControllerOverride;
      }
      set
      {
        _actionControllerOverride = value;
      }
    }

    public IActionController ActionController => ActionControllerOverride.Value;

    #region IActionControl implementation

    public event EventHandler<WorkbookActionEventArgs> ActionPerformed
    {
      add
      {
        if (this.ActionController != null)
        {
          this.ActionController.ActionPerformed += value;
        }
      }
      remove
      {
        if (this.ActionController != null)
        {
          this.ActionController.ActionPerformed -= value;
        }
      }
    }

    public event EventHandler<WorkbookActionEventArgs> Undid
    {
      add
      {
        if (this.ActionController != null)
        {
          this.ActionController.Undid += value;
        }
      }
      remove
      {
        if (this.ActionController != null)
        {
          this.ActionController.Undid -= value;
        }
      }
    }

    public event EventHandler<WorkbookActionEventArgs> Redid
    {
      add
      {
        if (this.ActionController != null)
        {
          this.ActionController.Redid += value;
        }
      }
      remove
      {
        if (this.ActionController != null)
        {
          this.ActionController.Redid -= value;
        }
      }
    }

    public void DoAction(Worksheet sheet, BaseWorksheetAction action) { this.ActionController.DoAction(sheet, action); }

    public void Undo() { this.ActionController.Undo(); }

    public void Redo() { this.ActionController.Redo(); }

    public void RepeatLastAction(RangePosition range) { this.ActionController.RepeatLastAction(range); }

    public void ClearActionHistory() { this.ActionController.ClearActionHistory(); }

    public void ClearActionHistoryForWorksheet(Worksheet sheet) { this.ActionController.ClearActionHistoryForWorksheet(sheet); }

    #endregion

    #endregion // Actions

    #region Settings

    /// <summary>
    /// Set specified workbook settings
    /// </summary>
    /// <param name="settings">Settings to be set</param>
    /// <param name="value">True to enable the settings, false to disable the settings</param>
    public void SetSettings(WorkbookSettings settings, bool value)
    {
      this.workbook.SetSettings(settings, value);
    }

    /// <summary>
    /// Get current settings of workbook
    /// </summary>
    /// <returns>Workbook settings set</returns>
    public WorkbookSettings GetSettings()
    {
      return this.workbook.GetSettings();
    }

    /// <summary>
    /// Determine whether or not the specified workbook settings has been set
    /// </summary>
    /// <param name="settings">Settings to be checked</param>
    /// <returns>True if specified settings has been set</returns>
    public bool HasSettings(WorkbookSettings settings)
    {
      return this.workbook.HasSettings(settings);
    }

    /// <summary>
    /// Enable specified settings for workbook.
    /// </summary>
    /// <param name="settings">Settings to be enabled.</param>
    public void EnableSettings(WorkbookSettings settings)
    {
      this.workbook.SetSettings(settings, true);
    }

    /// <summary>
    /// Disable specified settings for workbook.
    /// </summary>
    /// <param name="settings">Settings to be disabled.</param>
    public void DisableSettings(WorkbookSettings settings)
    {
      this.workbook.SetSettings(settings, false);
    }

    /// <summary>
    /// Event raised when settings is changed
    /// </summary>
    public event EventHandler SettingsChanged;
    #endregion // Settings

    #region Script

    /// <summary>
    /// Get or set script content
    /// </summary>
    public string Script
    {
      get { return this.workbook.Script; }
      set { this.workbook.Script = value; }
    }

#if EX_SCRIPT
			// TODO: srm should have only one instance 
			[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
			[Browsable(false)]
			public unvell.ReoScript.ScriptRunningMachine Srm
			{
				get { return this.workbook.Srm; }
			}

			/// <summary>
			/// Run workbook script.
			/// </summary>
			/// <returns>Return value from script.</returns>
			public object RunScript()
			{
				return this.workbook.RunScript();
			}

			/// <summary>
			/// Run specified script by workbook.
			/// </summary>
			/// <param name="script">Script to be executed.</param>
			/// <returns>Return value from specified script.</returns>
			public object RunScript(string script = null)
			{
				return this.workbook.RunScript(script);
			}
#endif

    #endregion // Script

    #region Internal Exceptions
    /// <summary>
    /// Event raised when exception has been happened during internal operations.
    /// Usually the internal operations are raised by hot-keys pressed by end-user.
    /// </summary>
    public event EventHandler<ExceptionHappenEventArgs> ExceptionHappened;

    void Workbook_ErrorHappened(object sender, ExceptionHappenEventArgs e)
    {
      this.ExceptionHappened?.Invoke(this, e);
    }

    /// <summary>
    /// Notify that there are exceptions happen on any worksheet. 
    /// The event ExceptionHappened of workbook will be invoked.
    /// </summary>
    /// <param name="sheet">Worksheet where the exception happened.</param>
    /// <param name="ex">Exception to describe the details of error information.</param>
    public void NotifyExceptionHappen(Worksheet sheet, Exception ex)
    {
      if (this.workbook != null)
      {
        this.workbook.NotifyExceptionHappen(sheet, ex);
      }
    }
    #endregion // Internal Exceptions

    #region Cursors
#if WINFORM || WPF
    private Cursor builtInCellsSelectionCursor = null;
    private Cursor builtInFullColSelectCursor = null;
    private Cursor builtInFullRowSelectCursor = null;
    private Cursor builtInEntireSheetSelectCursor = null;
    private Cursor builtInCrossCursor = null;

    private Cursor customCellsSelectionCursor = null;
    private Cursor defaultPickRangeCursor = null;
    private Cursor internalCurrentCursor = null;

    /// <summary>
    /// Get or set the mouse cursor on cells selection
    /// </summary>
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Cursor CellsSelectionCursor
    {
      get { return this.customCellsSelectionCursor ?? this.builtInCellsSelectionCursor; }
      set
      {
        this.customCellsSelectionCursor = value;
        this.internalCurrentCursor = value;
      }
    }

    /// <summary>
    /// Cursor symbol displayed when moving mouse over on row headers
    /// </summary>
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Cursor FullRowSelectionCursor { get; set; }

    /// <summary>
    /// Cursor symbol displayed when moving mouse over on column headers
    /// </summary>
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Cursor FullColumnSelectionCursor { get; set; }

    /// <summary>
    /// Get or set the mouse cursor of lead header part
    /// </summary>
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Cursor EntireSheetSelectionCursor { get; set; }

    private static Cursor LoadCursorFromResource(byte[] res)
    {
      using (var ms = new MemoryStream(res))
      {
        return new Cursor(ms);
      }
    }
#endif // WINFORM || WPF
    #endregion Cursors

    #region Pick Range
#if WINFORM || WPF
    /// <summary>
    /// Start to pick a range from current worksheet.
    /// </summary>
    /// <param name="onPicked">Callback function invoked after range is picked.</param>
    public void PickRange(Func<Worksheet, RangePosition, bool> onPicked)
    {
      this.PickRange(onPicked, this.defaultPickRangeCursor);
    }

    /// <summary>
    /// Start to pick a range from current worksheet.
    /// </summary>
    /// <param name="onPicked">Callback function invoked after range is picked.</param>
    /// <param name="pickerCursor">Cursor style during picking.</param>
    public void PickRange(Func<Worksheet, RangePosition, bool> onPicked, Cursor pickerCursor)
    {
      this.internalCurrentCursor = pickerCursor;

      this.currentWorksheet.PickRange((sheet, range) =>
      {
        bool ret = onPicked(sheet, range);
        return ret;
      });
    }

    /// <summary>
    /// Start to pick ranges and copy the styles to the picked range
    /// </summary>
    public void StartPickRangeAndCopyStyle()
    {
      this.currentWorksheet.StartPickRangeAndCopyStyle();
    }

    /// <summary>
    /// End pick range operation
    /// </summary>
    public void EndPickRange()
    {
      this.currentWorksheet.EndPickRange();

      this.internalCurrentCursor = (this.customCellsSelectionCursor ?? this.builtInCellsSelectionCursor);
    }
#endif // WINFORM || WPF
    #endregion // Pick Range

    #region Appearance

    /// <summary>
    /// Retrieve control instance of workbook.
    /// </summary>
    public ReoGridControl ControlInstance { get { return null; } }

    private ControlAppearanceStyle controlStyle = null;

    /// <summary>
    /// Control Style Settings
    /// </summary>
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    public ControlAppearanceStyle ControlStyle
    {
      get { return this.controlStyle; }
      set
      {
        if (value == null)
        {
          throw new ArgumentNullException("ControlStyle", "cannot set ControlStyle to null");
        }

        if (this.controlStyle != value)
        {
          if (this.controlStyle != null) this.controlStyle.CurrentControl = null;
          this.controlStyle = value;
        }
        //workbook.SetControlStyle(value);

        this.ApplyControlStyle();
      }
    }

    internal void ApplyControlStyle()
    {
      this.controlStyle.CurrentControl = this;

#if WINFORM
				sheetTab.BackColor = this.controlStyle[ControlAppearanceColors.SheetTabBackground];
				this.BackColor = this.controlStyle[ControlAppearanceColors.GridBackground];
#elif WPF
      sheetTab.Background = new System.Windows.Media.SolidColorBrush(this.controlStyle[ControlAppearanceColors.SheetTabBackground]);
#endif // WINFORM & WPF

      this.adapter?.Invalidate();
    }

    //private AppearanceStyle appearanceStyle = new AppearanceStyle(this);

    #endregion // Appearance

    #region Mouse
    private void OnWorksheetMouseDown(RGPointF location, MouseButtons buttons)
    {
      var sheet = this.currentWorksheet;

      if (sheet != null)
      {
        // if currently control is in editing mode, make the input fields invisible
        if (sheet.currentEditingCell != null)
        {
          if (this.adapter is IEditableControlAdapter editableAdapter)
          {
            sheet.EndEdit(editableAdapter.GetEditControlText());
          }
        }

        sheet.ViewportController?.OnMouseDown(location, buttons);
      }
    }

    private void OnWorksheetMouseMove(RGPointF location, MouseButtons buttons)
    {
      this.currentWorksheet?.ViewportController?.OnMouseMove(location, buttons);
    }

    private void OnWorksheetMouseUp(RGPointF location, MouseButtons buttons)
    {
      this.currentWorksheet?.ViewportController?.OnMouseUp(location, buttons);
    }
    #endregion // Mouse

#if WINFORM || WPF
#if WINFORM
			/// <summary>
			/// Overrides mouse-leave event
			/// </summary>
			/// <param name="e">Argument of mouse-leave</param>
			protected override void OnMouseLeave(EventArgs e)
			{
#elif WPF
    protected override void OnMouseLeave(System.Windows.Input.MouseEventArgs e)
    {
#endif // WPF
      base.OnMouseLeave(e);

      if (this.currentWorksheet != null)
      {
        this.adapter.ChangeCursor(CursorStyle.PlatformDefault);
        this.currentWorksheet.HoverPos = CellPosition.Empty;
      }
    }
#endif // WINFORM || WPF

#if PRINT
    /// <summary>
    /// Create a print session to print all worksheets.
    /// </summary>
    /// <returns>Print session to print specified worksheets.</returns>
    public Print.PrintSession CreatePrintSession()
    {
      return this.workbook.CreatePrintSession();
    }
#endif // PRINT

    #region SheetTabControl

    /// <summary>
    /// Show or hide the built-in sheet tab control.
    /// </summary>
    public bool SheetTabVisible
    {
      get { return (this.HasSettings(WorkbookSettings.View_ShowSheetTabControl)); }
      set
      {
        if (value)
        {
          this.EnableSettings(WorkbookSettings.View_ShowSheetTabControl);
        }
        else
        {
          this.DisableSettings(WorkbookSettings.View_ShowSheetTabControl);
        }
      }
    }

    /// <summary>
    /// Get or set the width of sheet tab control.
    /// </summary>
    public IntOrDouble SheetTabWidth
    {
      get { return this.sheetTab.ControlWidth; }
      set { this.sheetTab.ControlWidth = value; }
    }

    /// <summary>
    /// Determines that whether or not to display the new button on sheet tab control.
    /// </summary>
    public bool SheetTabNewButtonVisible
    {
      get { return this.sheetTab.NewButtonVisible; }
      set { this.sheetTab.NewButtonVisible = value; }
    }
    #endregion // SheetTabControl

    #region Scroll

    /// <summary>
    /// Scroll current active worksheet.
    /// </summary>
    /// <param name="x">Scroll value on horizontal direction.</param>
    /// <param name="y">Scroll value on vertical direction.</param>
    public void ScrollCurrentWorksheet(RGFloat x, RGFloat y)
    {
      if (this.currentWorksheet?.ViewportController is IScrollableViewportController svc)
      {
        svc.ScrollViews(ScrollDirection.Both, x, y);

        svc.SynchronizeScrollBar();
      }
    }

    /// <summary>
    /// Event raised when current worksheet is scrolled.
    /// </summary>
    public event EventHandler<WorksheetScrolledEventArgs> WorksheetScrolled;

    /// <summary>
    /// Raise the event of worksheet scrolled.
    /// </summary>
    /// <param name="worksheet">Instance of scrolled worksheet.</param>
    /// <param name="x">Scroll value on horizontal direction.</param>
    /// <param name="y">Scroll value on vertical direction.</param>
    public void RaiseWorksheetScrolledEvent(Worksheet worksheet, RGFloat x, RGFloat y)
    {
      this.WorksheetScrolled?.Invoke(this, new WorksheetScrolledEventArgs(worksheet)
      {
        X = x,
        Y = y,
      });
    }

    private bool showScrollEndSpacing = true;

    [DefaultValue(100)]
    [Browsable(true)]
    [Description("Determines whether or not show the white spacing at bottom and right of worksheet.")]
    public bool ShowScrollEndSpacing
    {
      get { return this.showScrollEndSpacing; }
      set
      {
        if (this.showScrollEndSpacing != value)
        {
          this.showScrollEndSpacing = value;
          this.currentWorksheet.UpdateViewportController();
        }
      }
    }

    #endregion Scroll
  }
}
