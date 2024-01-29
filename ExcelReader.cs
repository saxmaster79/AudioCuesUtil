
using SpreadsheetLight;

namespace AudioCuesUtil;
    public class ExcelReader
    {
    internal Show ReadShow(FileInfo excelInputFile)
    {
        SLDocument sl = new SLDocument(excelInputFile.FullName, "Show");

        SLWorksheetStatistics stats = sl.GetWorksheetStatistics();
        var propertyNames = new List<string>() { "" };
        for (int i = 1; i <= stats.EndColumnIndex; i++)
        {
            propertyNames.Add(sl.GetCellValueAsString(1, i));
        }

        string showTitle = sl.GetCellValueAsString(2, propertyNames.IndexOf(nameof(Show.ShowTitle)));
        string showId = sl.GetCellValueAsString(2, propertyNames.IndexOf(nameof(Show.ShowId)));

        sl.SelectWorksheet("Cues");
        var cues = ReadRows(sl).ToList();

        return new Show(cues, showTitle, showId);
    }

    private IEnumerable<Cue> ReadRows(SLDocument sl) 
        {
            var propertyNames = new List<string>() { "" };
            SLWorksheetStatistics stats = sl.GetWorksheetStatistics();
            for (int row = 1; row <= stats.EndColumnIndex; row++)
            {
                propertyNames.Add(sl.GetCellValueAsString(1, row));
            }
            var type = typeof(Cue);
            for (int row = 2; row <= stats.EndRowIndex; row++)
            {
                var rowObject = new Cue();
                for (int col = 1; col <= stats.EndColumnIndex; col++)
                {
                    var prop = type.GetProperty(propertyNames[col]);
                    if (prop == null) continue;
                    object value;
                    var stringValue = sl.GetCellValueAsString(row, col);
                    switch (prop.PropertyType)
                    {
                        case Type _ when prop.PropertyType == typeof(string):
                            value = stringValue;
                            break;
                        case Type _ when prop.PropertyType == typeof(int):
                            if (string.IsNullOrEmpty(stringValue)) continue;
                            value = int.Parse(stringValue);
                            break;
                        case Type _ when prop.PropertyType == typeof(float):
                            if (string.IsNullOrEmpty(stringValue)) continue;
                            value = float.Parse(stringValue);
                            break;
                        default: throw new ArgumentException(prop.PropertyType + " unknown");
                    }
                    prop.SetValue(rowObject, value, null);


                }
                if (!string.IsNullOrEmpty(rowObject.CueId) )
                    yield return rowObject;
            }

        }
    }
