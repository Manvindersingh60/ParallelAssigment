using ExcelDna.Integration;

namespace ParallelAssigment
{
    /// <summary>
    /// Contains Excel UDFs
    /// </summary>
    public static class Functions
    {
        /// <summary>
        /// Increments number by 2
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        [ExcelFunction(Description ="Adds 2 to the number")]
        public static object AddTwo(int num)
        {
            var caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;  //used for stoping caching
            return ExcelAsyncUtil.Run("AddTwo", new object[] {num, caller }, () => Utility.SendRequest(num));
        }
        
    }
}
