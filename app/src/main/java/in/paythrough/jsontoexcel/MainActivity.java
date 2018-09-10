package in.paythrough.jsontoexcel;

import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.widget.Toast;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import java.io.IOException;

import static in.paythrough.jsonarraytoexcelcsv.JSONArrayToExcel.AssetJSONFile;
import static in.paythrough.jsonarraytoexcelcsv.JSONArrayToExcel.saveExcelFile;


public class MainActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        try {
            String jsonString = AssetJSONFile("tatkal.json", MainActivity.this);
            JSONObject jsonObject = new JSONObject(jsonString);
            String responseCode = jsonObject.getString("responseCode");
            if (responseCode.equals("200")) {
                JSONArray jsonArray = jsonObject.getJSONArray("withdrawReport");
                if(jsonArray.length()>0){
                    if(saveExcelFile(MainActivity.this,"MyExcel",jsonArray,true)){
                        Toast.makeText(this, "SUCCESS", Toast.LENGTH_SHORT).show();
                    }else{
                        Toast.makeText(this, "NOT SUCCESS", Toast.LENGTH_SHORT).show();
                    }
                }else{
                    Toast.makeText(this, "No Data Found", Toast.LENGTH_SHORT).show();
                    return;
                }
            }else{
                Toast.makeText(this, "Something Wrong", Toast.LENGTH_SHORT).show();
                return;
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (JSONException e) {
            e.printStackTrace();
        }
    }
}