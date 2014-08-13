using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/// <summary>
/// Summary description for UtilityManager
/// </summary>
public class UtilityManager
{
	//Constructor
    public UtilityManager(){
	}
    
    public static int GetNextIndex(string[,] array) {
        if (array[0, 0] == null) { return 0; }

        int row = 0;
        for (int x = 0; x < array.GetLength(0); x++) {
            if (array[x, 0] == null) {
                break;
            }
            row = x;
        }
        return row + 1;
    }
    
    public static string[,] CompactThisTwoDimensionalArray(string[,] m) {
        int noOfCols = m.Length / Convert.ToInt32(m.GetLongLength(0));

        int noOfRows = 0;
        for (int i = 0; i < Convert.ToInt32(m.GetLongLength(0)); i++) {
            if (m.GetValue(i, 0) != null) {
                noOfRows++;
            }
        }

        string[,] n = new string[noOfRows, noOfCols];

        int row = -1;
        for (int i = 0; i < Convert.ToInt32(m.GetLongLength(0)); i++) {
            if (m.GetValue(i, 0) != null) {
                row++;
                for (int x = 0; x < noOfCols; x++) {
                    n[row, x] = m[i, x];
                }
            }
        }
        return n;
    }
    
    public static string[,] ResetThisTwoDimensionalArray(string[,] m) {
        for (int row = 0; row < m.GetLength(0); row++) {
            for (int col = 0; col < m.GetLength(1); col++) {
                m[row, col] = null;
            }
        }
        return m;
    }
		
}
