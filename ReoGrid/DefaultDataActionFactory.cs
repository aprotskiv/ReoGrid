using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Core;
using unvell.ReoGrid.Interaction;

namespace unvell.ReoGrid
{
  public interface IDataActionFactory
  {
    /// <summary>
    /// 
    /// </summary>
    /// <param name="range"></param>
    /// <param name="partialGrid"></param>
    /// <returns></returns>
    CutRangeAction CutRangeAction(RangePosition range, PartialGrid partialGrid);

    /// <summary>
    /// Create action to remove data from specified range.
    /// </summary>
    /// <param name="range">data from cells in this range will be removed.</param>
    /// <param name="keyData">pressed keys that triggered <see cref="RemoveRangeDataAction"/></param>
    RemoveRangeDataAction RemoveRangeDataAction(RangePosition range, KeyCode? keyData = null);
    
    

    /// <summary>
    /// Create action to set partial grid.
    /// </summary>
    /// <param name="range">target range to set partial grid.</param>
    /// <param name="data">partial grid to be set.</param>
    WorksheetReusableAction SetPartialGridAction(RangePosition range, PartialGrid data);

    /// <summary>
    /// Create action to set data into specified range of spreadsheet.
    /// </summary>
    /// <param name="range">range to set specified data.</param>
    /// <param name="matrix">object[,] to be set.</param>
    WorksheetReusableAction SetRangeDataAction(RangePosition range, SetRangeDataActionContext matrix);


    BaseWorksheetAction SetCellDataAction(int internalRow, int internalCol, object data);
  }

  public sealed class SetRangeDataActionContext
  {
    public SetRangeDataActionContext(object[,] arrayData)
    {
      this.Data = arrayData;
    }

    /// <summary>
    /// data to be set
    /// </summary>
    public object[,] Data { get; }
  }

  /// <summary>
  /// Creates <see cref="WorksheetReusableAction"/> instances
  /// </summary>
  public class DefaultDataActionFactory : IDataActionFactory
  {
    /// <inheritdoc/>
    public CutRangeAction CutRangeAction(RangePosition range, PartialGrid partialGrid)
    {
      return new CutRangeAction(range, partialGrid);
    }

    /// <inheritdoc/>
    public RemoveRangeDataAction RemoveRangeDataAction(RangePosition range, KeyCode? keyData = null)
    {
      return new RemoveRangeDataAction(range, keyData);
    }

    /// <inheritdoc/>
    public WorksheetReusableAction SetPartialGridAction(RangePosition range, PartialGrid data)
    {
      return new SetPartialGridAction(range, data);
    }
    
    /// <inheritdoc/>
    public WorksheetReusableAction SetRangeDataAction(RangePosition range, SetRangeDataActionContext data)
    {
      return new SetRangeDataAction(range, data.Data);
    }

    public BaseWorksheetAction SetCellDataAction(int internalRow, int internalCol, object data)
    {
      return new SetCellDataAction(internalRow, internalCol, data);
    }
  }
}
