namespace unvell.ReoGrid
{
  public interface IPartialGridFactory
  {
    PartialGrid GetPartialGrid(Worksheet worksheet,
        RangePosition range, PartialGridCopyFlag flag, ExPartialGridCopyFlag exFlag, bool checkIntersectedRange
    );
  }

  public class DefaultPartialGridFactory : IPartialGridFactory
  {
    public PartialGrid GetPartialGrid(Worksheet worksheet,
        RangePosition range, PartialGridCopyFlag flag, ExPartialGridCopyFlag exFlag, bool checkIntersectedRange
      )
    {
      return worksheet.GetPartialGrid(range, flag, exFlag, checkIntersectedRange);
    }
  }
}
