package in.paythrough.jsonarraytoexcelcsv;

import android.app.ProgressDialog;
import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

public class JSONArrayToExcel {

    public static boolean saveExcelFile(Context context, String fileName, JSONArray jsonArray, boolean showProgressDialog) {

        ProgressDialog dialog;
        dialog = new ProgressDialog(context);
        dialog.setMessage("Wait Until Download. Download Folder is your internal download folder");
        dialog.setCancelable(false);
        dialog.setTitle(fileName+".csv file Download");

        if(showProgressDialog) dialog.show();

        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            if(showProgressDialog) dialog.dismiss();
            Toast.makeText(context, "EXTERNAL STORAGE PERMISSION ERROR", Toast.LENGTH_SHORT).show();
            return false;
        }
        if (jsonArray.length()<=0) {
            if(showProgressDialog) dialog.dismiss();
            Toast.makeText(context, "Null Value", Toast.LENGTH_SHORT).show();
            return false;
        }

        boolean success = false;

        //New Workbook
        Workbook wb = new HSSFWorkbook();

        Cell c = null;

        //Cell style for header row
        CellStyle cs = wb.createCellStyle();

        cs.setFillForegroundColor(HSSFColor.AQUA.index);

        cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        //New Sheet
        Sheet sheet1 = null;
        sheet1 = wb.createSheet("PAYTHROUGH");

        // Generate column headings
        org.apache.poi.ss.usermodel.Row row = sheet1.createRow(0);

        try{
            //CREATE HEADER PART
            JSONObject jsonObject = jsonArray.getJSONObject(0);
            Iterator<String> iter = jsonObject.keys();
            int i = 0;
            while (iter.hasNext()) {
                String key = iter.next();

                c = row.createCell(i);
                c.setCellValue(key);
                c.setCellStyle(cs);

                sheet1.setColumnWidth(i, (15 * 400));
                i++;
            }


            int j = 1;
            for(int k=0; k<jsonArray.length(); k++){
                JSONObject jsonObjectData = jsonArray.getJSONObject(k);
                Iterator<String> iterData = jsonObjectData.keys();

                JSONObject jsonObjectDataPart = jsonArray.getJSONObject(k);
                org.apache.poi.ss.usermodel.Row rowData = sheet1.createRow(j);

                int m = 0;
                while (iterData.hasNext()) {
                    String key = iterData.next();

                    //String dataValue = jsonObjectDataPart.getString(key);
                    String dataValue = checkKeyAndGetValue(jsonObjectDataPart,key);

                    c = rowData.createCell(m);
                    c.setCellValue(dataValue);

                    //sheet1.setColumnWidth(m, (15 * 400));
                    m++;
                }
                j++;
            }
        }catch (Exception e){
            if(showProgressDialog) dialog.dismiss();
            e.printStackTrace();
        }

        // Create a path where we will place our List of objects on external storage
        //File file = new File(context.getExternalFilesDir(null), fileName);
        //File file = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), fileName + ".csv");
        File file = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), fileName + ".xlsx");
        FileOutputStream os = null;

        try {
            os = new FileOutputStream(file);
            wb.write(os);
            Log.w("FileUtils", "Writing file" + file);
            success = true;
        } catch (IOException e) {
            if(showProgressDialog) dialog.dismiss();
            Log.w("FileUtils", "Error writing " + file, e);
        } catch (Exception e) {
            if(showProgressDialog) dialog.dismiss();
            Log.w("FileUtils", "Failed to save file", e);
        } finally {
            try {
                if (null != os)
                    os.close();
            } catch (Exception ex) {
            }
        }
        if(showProgressDialog) dialog.dismiss();
        return success;
    }



    public static boolean isExternalStorageReadOnly() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState)) {
            return true;
        }
        return false;
    }

    public static boolean isExternalStorageAvailable() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(extStorageState)) {
            return true;
        }
        return false;
    }

    public static String AssetJSONFile (String filename, Context context) throws IOException {
        AssetManager manager = context.getAssets();
        InputStream file = manager.open(filename);
        byte[] formArray = new byte[file.available()];
        file.read(formArray);
        file.close();

        return new String(formArray);
    }

    public static String checkKeyAndGetValue(JSONObject jsonObject, String key){
        String value ="";
        try{
            value = jsonObject.has(key)?jsonObject.getString(key):"";
        }catch (Exception e){
            e.printStackTrace();
        }
        return value;
    }
}
