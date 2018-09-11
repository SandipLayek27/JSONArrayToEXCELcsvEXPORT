# JSONArrayToEXCELcsvEXPORT
It is a intelligent library system where most immportant inputed field is JSONArray it atometically detect Header names and re-arrange those names with particular valid data which hold JSONObject data.
JSONArray to excel (.csv) export system.
We create this library for my working purpose.
Here, we used some pre generated libraries and modify it's working process.

## Developed
[![Sandip](https://avatars1.githubusercontent.com/u/31722942?v=4&u=18643bfaaba26114584d27693e9891db26bcb582&s=39) Sandip](https://github.com/SandipLayek27)  
# ★ Gradle Dependency
Add Gradle dependency in the build.gradle file of your application module (app in the most cases) :
First Tab:

```sh
allprojects{
    repositories{
        jcenter()
        maven { url 'https://jitpack.io' }
    }
}
```

AND

```sh
dependencies {
    implementation 'com.github.SandipLayek27:JSONArrayToEXCELcsvEXPORT:1.1'
}
```

# ★ Features are
1. JSONArray to Excel Converter(.csv format).


# ★ Uses
```sh
❆ 1 PERMISSIONS:-
<uses-permission android:name="android.permission.READ_EXTERNAL_STORAGE" />
<uses-permission android:name="android.permission.WRITE_EXTERNAL_STORAGE" />

❆ 2 CODE:-
        try {
            String jsonString = AssetJSONFile("jsonfile.json", MainActivity.this);
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
        
❆ 2 NOTES:-
Here we already paste a json formated file to my project asset folder as example purpose.
saveExcelFile(MainActivity.this,"MyExcel",jsonArray,true)
// MainActivity.this => context
// MyExcel => Excel file name
// jsonArray => JSONArray 
// true => If you want to show progerss bar then put true otherwise false.
// Download .csv file to your internal download folder.        
```

