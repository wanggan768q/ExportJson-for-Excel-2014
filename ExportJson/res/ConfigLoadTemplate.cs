using UnityEngine;
using System.Collections;
using HS.Base;

public class ConfigLoad : HS_SingletonGameObject<ConfigLoad> {

	private string textContent;

	private int fileCount = $fileCount$;

	public delegate void ConfigLoadProgress(float f);

    public event ConfigLoadProgress configLoadProgress;

	public IEnumerator LoadConfig () {

$loadConfItem$

		yield return true;
	}

    IEnumerator LoadData (string name) {

		string path = HS_Base.GetStreamingAssetsFilePath(name, "json");
	
		WWW www = new WWW(path);
		yield return www;

		textContent = www.text;
		yield return true;
	}

	void Progress(int index)
    {
        if (configLoadProgress != null)
        {
            configLoadProgress((float)index / (float)fileCount);
        }
    }
}
