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

using unvell.Common;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Main;

namespace unvell.ReoGrid
{
  public class ActionController : IActionController
  {
    protected readonly ICurrentWorksheetProvider _currentWorksheetProvider;
    protected readonly ActionManager actionManager;
    protected WorksheetReusableAction lastReusableAction;

    protected Worksheet currentWorksheet => _currentWorksheetProvider.CurrentWorksheet;
    protected Worksheet CurrentWorksheet
    {
      set
      {
        _currentWorksheetProvider.CurrentWorksheet = value;
      }
    }

    public ActionController(ICurrentWorksheetProvider currentWorksheetProvider)
    {
      _currentWorksheetProvider = currentWorksheetProvider;
      actionManager = new ActionManager();

      this.actionManager.BeforePerformAction += (s, e) =>
      {
        if (this.BeforeActionPerform != null)
        {
          var arg = new BeforeActionPerformEventArgs(e.Action);
          this.BeforeActionPerform(this, arg);
          e.Cancel = arg.IsCancelled;
        }
      };

      // register for moniting reusable action
      this.actionManager.AfterPerformAction += (s, e) =>
      {
        if (e.Action is WorksheetReusableAction)
        {
          this.lastReusableAction = e.Action as WorksheetReusableAction;
        }

        this.ActionPerformed?.Invoke(this, new WorkbookActionEventArgs(e.Action));
      };
    }

    public void DoAction(BaseWorksheetAction action)
    {
      this.DoAction(this.currentWorksheet, action);
    }

    /// <summary>Do specified action. 
    /// 
    /// An action does the operation as well as undoes for worksheet.
    /// Actions performed by this method will be appended to action history stack 
    /// in order to undo, redo and repeat.
    /// 
    /// There are built-in actions available for many base operations, such as:
    ///   <code>SetCellDataAction</code> - set cell data
    ///   <code>SetRangeDataAction</code> - set data into range
    ///   <code>SetRangeBorderAction</code> - set border to specified range
    ///   <code>SetRangeStyleAction</code> - set styles to specified range
    ///   ...
    ///   
    /// It is possible to make custom action by inherting BaseWorksheetAction.
    /// </summary>
    /// <example>
    /// ReoGrid uses ActionManager, unvell lightweight undo framework, 
    /// to implement the Do/Undo/Redo/Repeat method.
    /// 
    /// To do action:
    /// <code>
    ///   var action = new SetCellDataAction("B1", 10);
    ///   workbook.DoAction(targetSheet, action);
    /// </code>
    /// 
    /// To undo action:
    /// <code>
    ///   workbook.Undo();
    /// </code>
    /// 
    /// To redo action:
    /// <code>
    ///			workbook.Redo();
    /// </code>
    /// 
    /// To repeat last action:
    /// <code>
    ///			workbook.RepeatLastAction(targetSheet, new ReoGridRange("B1:C3"));
    /// </code>
    /// 
    /// It is possible to do multiple actions at same time:
    /// <code>
    ///   var action1 = new SetRangeDataAction(...);
    ///   var action2 = new SetRangeBorderAction(...);
    ///   var action3 = new SetRangeStyleAction(...);
    ///   
    ///			var actionGroup = new WorksheetActionGroup();
    ///			actionGroup.Actions.Add(action1);
    ///			actionGroup.Actions.Add(action2);
    ///			actionGroup.Actions.Add(action3);
    ///			
    ///			workbook.DoAction(targetSheet, actionGroup);
    /// </code>
    /// 
    /// Actions added into action group will be performed by one time,
    /// they will be also undone by one time.
    /// </example>
    /// <seealso cref="ActionGroup"/>
    /// <seealso cref="BaseWorksheetAction"/>
    /// <seealso cref="WorksheetActionGroup"/>
    /// <param name="sheet">worksheet of the target container to perform specified action</param>
    /// <param name="action">action to be performed</param>
    public virtual void DoAction(Worksheet sheet, BaseWorksheetAction action)
    {
      action.Worksheet = sheet;

      this.actionManager.DoAction(action);

      if (action is WorksheetReusableAction reusableAction)
      {
        this.lastReusableAction = reusableAction;
      }

      if (this.currentWorksheet != sheet)
      {
        sheet.RequestInvalidate();
        this.CurrentWorksheet = sheet;
      }

      // fix #282, https://github.com/unvell/ReoGrid/issues/282
      // comment out to avoid invoke ActionPerformed event, which is already invoked by actionManager above.
      //if (ActionPerformed != null) ActionPerformed(this, new WorkbookActionEventArgs(action));
    }

    /// <summary>
    /// Undo the last action.
    /// </summary>
    public virtual void Undo()
    {
      if (this.currentWorksheet != null)
      {
        if (this.currentWorksheet.IsEditing)
        {
          this.currentWorksheet.EndEdit(EndEditReason.NormalFinish);
        }
      }

      var action = this.actionManager.Undo();

      if (action != null)
      {
        if (action is WorkbookAction)
        {
          // seems nothing to do
        }
        else if (action is BaseWorksheetAction worksheetAction)
        {
          var sheet = worksheetAction.Worksheet;

          if (action is WorksheetReusableAction reusableAction)
          {
            if (sheet != null)
            {
              sheet.SelectRange(reusableAction.Range);
            }
          }

          if (sheet != null)
          {
            sheet.RequestInvalidate();
            this.CurrentWorksheet = sheet;
          }
        }

        RaiseUndid(new WorkbookActionEventArgs(action));        
      }
    }

    /// <summary>
    /// Redo the last action.
    /// </summary>
    public virtual void Redo()
    {
      if (this.currentWorksheet != null)
      {
        if (this.currentWorksheet.IsEditing)
        {
          this.currentWorksheet.EndEdit(EndEditReason.NormalFinish);
        }
      }

      var action = this.actionManager.Redo();

      if (action != null)
      {
        if (action is BaseWorksheetAction worksheetAction)
        {
          var sheet = worksheetAction.Worksheet;

          if (action is WorksheetReusableAction reusableAction)
          {
            this.lastReusableAction = reusableAction;

            if (sheet != null)
            {
              sheet.SelectRange(this.lastReusableAction.Range);
            }
          }

          if (sheet != null && this.currentWorksheet != sheet)
          {
            sheet.RequestInvalidate();
            this.CurrentWorksheet = sheet;
          }
        }

        RaiseRedid(new WorkbookActionEventArgs(action));
      }
    }

    /// <summary>
    /// Repeat to do last action and apply to another specified range.
    /// </summary>
    /// <param name="range">The new range to be applied for the last action.</param>
    public void RepeatLastAction(RangePosition range)
    {
      this.RepeatLastAction(this.currentWorksheet, range);
    }

    /// <summary>
    /// Repeat to do last action and apply to another specified range and worksheet.
    /// </summary>
    /// <param name="worksheet">The target worksheet to perform the action.</param>
    /// <param name="range">The new range to be applied for the last action.</param>
    public virtual void RepeatLastAction(Worksheet worksheet, RangePosition range)
    {
      if (this.currentWorksheet != null)
      {
        if (this.currentWorksheet.IsEditing)
        {
          this.currentWorksheet.EndEdit(EndEditReason.NormalFinish);
        }
      }

      if (this.CanRedo())
      {
        this.Redo();
      }
      else
      {
        if (this.lastReusableAction != null)
        {
          var newAction = lastReusableAction.Clone(range);
          newAction.Worksheet = worksheet;

          this.actionManager.DoAction(newAction);

          // fix #282, https://github.com/unvell/ReoGrid/issues/282
          //this.ActionPerformed?.Invoke(this, new WorkbookActionEventArgs(newAction));

          this.currentWorksheet.RequestInvalidate();
        }
      }
    }

    /// <summary>
    /// Determine whether there is any actions can be undone.
    /// </summary>
    /// <returns>True if any actions can be undone</returns>
    public bool CanUndo()
    {
      return this.actionManager.CanUndo();
    }

    /// <summary>
    /// Determine whether there is any actions can be redid.
    /// </summary>
    /// <returns>True if any actions can be redid</returns>
    public bool CanRedo()
    {
      return this.actionManager.CanRedo();
    }

    /// <summary>
    /// Clear all undo/redo actions from workbook action history.
    /// </summary>
    public void ClearActionHistory()
    {
      this.actionManager.Reset();

      this.lastReusableAction = null;
    }

    /// <summary>
    /// Delete all actions that belongs to specified worksheet.
    /// </summary>
    /// <param name="sheet">Actions belongs to this worksheet will be deleted from workbook action histroy.</param>
    public virtual void ClearActionHistoryForWorksheet(Worksheet sheet)
    {
      List<IUndoableAction> undoActions = this.actionManager.UndoStack;
      for (int i = 0; i < undoActions.Count;)
      {
        var action = undoActions[i];

        var worksheetAction = (action as BaseWorksheetAction);

        if (worksheetAction != null && worksheetAction.Worksheet == sheet)
        {
          undoActions.RemoveAt(i);
          continue;
        }

        i++;
      }

      int totalActions = undoActions.Count;
      var redoActions = new List<IUndoableAction>(this.actionManager.RedoStack);

      for (int i = 0; i < redoActions.Count;)
      {
        IUndoableAction action = redoActions[i];

        var worksheetAction = (action as BaseWorksheetAction);

        if (worksheetAction != null && worksheetAction.Worksheet == sheet)
        {
          redoActions.RemoveAt(i);
          continue;
        }

        i++;
      }

      this.actionManager.RedoStack.Clear();

      for (int i = redoActions.Count - 1; i >= 0; i--)
      {
        this.actionManager.RedoStack.Push(redoActions[i]);
      }

      totalActions += redoActions.Count;

      if (totalActions <= 0)
      {
        this.lastReusableAction = null;
      }
    }

    /// <summary>
    /// Event fired before action perform.
    /// </summary>
    public event EventHandler<WorkbookActionEventArgs> BeforeActionPerform;

    /// <summary>
    /// Event fired when any action performed.
    /// </summary>
    public event EventHandler<WorkbookActionEventArgs> ActionPerformed;

    /// <summary>
    /// Event fired when Undo operation performed by user.
    /// </summary>
    public event EventHandler<WorkbookActionEventArgs> Undid;

    /// <summary>
    /// Event fired when Reod operation performed by user.
    /// </summary>
    public event EventHandler<WorkbookActionEventArgs> Redid;

    protected void RaiseRedid(WorkbookActionEventArgs workbookActionEventArgs)
    {
      Redid?.Invoke(this, workbookActionEventArgs);
    }

    protected void RaiseUndid(WorkbookActionEventArgs workbookActionEventArgs)
    {
      Undid?.Invoke(this, workbookActionEventArgs);
    }
  }
}