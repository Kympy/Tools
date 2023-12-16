
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.AddressableAssets;

[System.Serializable]
public class Data
{
    public int Id;
}

public class TableBase<T> where T : Data
{
    public static T[] FromJson(string json)
    {
        Wrapper wrapper = JsonUtility.FromJson<Wrapper>(json);
        return wrapper.Items;
    }
    // public static string ToJson(T[] array)
    // {
    //     Wrapper wrapper = new Wrapper();
    //     wrapper.Items = array;
    //     return JsonUtility.ToJson(wrapper);
    // }
    public static void LoadTable()
    {
        var handle = Addressables.LoadAssetAsync<TextAsset>($"{ResourcePath.Table}/{typeof(T)}.json");
        handle.WaitForCompletion();
        TextAsset jsonData = handle.Result;
        if (jsonData == null)
        {
            Debug.LogError($"{typeof(T)} table json data is null.");
            return;
        }
        T[] items = FromJson($"{{\"Items\":{jsonData.text}}}");
        table = new Dictionary<int, T>();
        table.Clear();
        for(int i = 0; i < items.Length; i++)
        {
            table.Add(items[i].Id, items[i]);           
        }
        Debug.Log($"{typeof(T)} table loaded.");
    }
    [System.Serializable]
    public class Wrapper
    {
        public T[] Items;
    }
    private static Dictionary<int, T> table;
    public static Dictionary<int, T> Table
    {
        get
        {
            return table;
        }
    }
};
