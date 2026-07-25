namespace unvell.ReoGrid
{
  public interface IPartialGridFactory
  {
    PartialGrid GetPartialGrid_ForCopyAction(Worksheet worksheet,
        RangePosition range, PartialGridCopyFlag flag, ExPartialGridCopyFlag exFlag, bool checkIntersectedRange
    );


    PartialGrid GetPartialGrid_ForRemoveRangeAction(Worksheet worksheet,
        RangePosition range, PartialGridCopyFlag flag, ExPartialGridCopyFlag exFlag, bool checkIntersectedRange
    );
  }

  public class DefaultPartialGridFactory : IPartialGridFactory
  {
    public PartialGrid GetPartialGrid_ForCopyAction(Worksheet worksheet,
        RangePosition range, PartialGridCopyFlag flag, ExPartialGridCopyFlag exFlag, bool checkIntersectedRange
      )
    {
      return worksheet.GetPartialGrid(range, flag, exFlag, checkIntersectedRange);
    }

    public PartialGrid GetPartialGrid_ForRemoveRangeAction(Worksheet worksheet,
        RangePosition range, PartialGridCopyFlag flag, ExPartialGridCopyFlag exFlag, bool checkIntersectedRange
      )
    {
      return worksheet.GetPartialGrid(range, flag, exFlag, checkIntersectedRange);
    }

  }
}
