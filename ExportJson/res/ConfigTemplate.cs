//郭晓波
using UnityEngine;
using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using HS.IO;
using LitJson;


public class $Template$Element
{
$FieldDefine$
	public bool IsValidate = false;
	public $Template$Element()
	{
$InitPrimaryField$
	}
};


public class $Template$Table
{

	private $Template$Table()
	{
		_MapElements = new Dictionary<int, $Template$Element>();
		_EmptyItem = new $Template$Element();
		_VecAllElements = new List<$Template$Element>();
	}
	private Dictionary<int, $Template$Element> _MapElements = null;
	private List<$Template$Element>	_VecAllElements = null;
	private $Template$Element _EmptyItem = null;
	private static $Template$Table _SInstance = null;

	public static $Template$Table Instance
	{
		get
		{
			if( _SInstance != null )
				return _SInstance;	
			_SInstance = new $Template$Table();
			return _SInstance;
		}
	}

	public $Template$Element GetElement(int key)
	{
		if( _MapElements.ContainsKey(key) )
			return _MapElements[key];
		return _EmptyItem;
	}

	public int GetElementCount()
	{
		return _MapElements.Count;
	}
	public bool HasElement(int key)
	{
		return _MapElements.ContainsKey(key);
	}

  public List<$Template$Element> GetAllElement(Predicate<$Template$Element> matchCB = null)
	{
        if( matchCB==null || _VecAllElements.Count == 0)
            return _VecAllElements;
        return _VecAllElements.FindAll(matchCB);
	}

	public bool Load()
	{
		
		string strTableContent = "";
		if(HS_ByteRead.ReadCsvFile("$Template$.json", out strTableContent ) )
			return LoadCsv( strTableContent );
		byte[] binTableContent = null;
		if( !HS_ByteRead.ReadBinFile("$Template$.bin", out binTableContent ) )
		{
			Debug.Log("配置文件[$Template$.bin]未找到");
			return false;
		}
		return LoadBin(binTableContent);
	}


	public bool LoadBin(byte[] binContent)
	{
		_MapElements.Clear();
		_VecAllElements.Clear();
		int nCol, nRow;
		int readPos = 0;
		readPos += HS_ByteRead.ReadInt32Variant( binContent, readPos, out nCol );
		readPos += HS_ByteRead.ReadInt32Variant( binContent, readPos, out nRow );
		List<string> vecLine = new List<string>(nCol);
		List<int> vecHeadType = new List<int>(nCol);
        string tmpStr;
        int tmpInt;
		for( int i=0; i<nCol; i++ )
		{
            readPos += HS_ByteRead.ReadString(binContent, readPos, out tmpStr);
            readPos += HS_ByteRead.ReadInt32Variant(binContent, readPos, out tmpInt);
            vecLine.Add(tmpStr);
            vecHeadType.Add(tmpInt);
		}
		if(vecLine.Count != $ColCount$)
		{
			Debug.Log("$Template$.json 中列数量与生成的代码不匹配!");
			return false;
		}
$CheckColName$
		for(int i=0; i<nRow; i++)
		{
			$Template$Element member = new $Template$Element();
$ReadBinColValue$
			member.IsValidate = true;
			_VecAllElements.Add(member);
			_MapElements[member.$PrimaryKey$] = member;
		}
		return true;
	}
	public bool LoadCsv(string strContent)
	{
		if( strContent.Length == 0 )
			return false;
		_MapElements.Clear();
		_VecAllElements.Clear();
		int contentOffset = 0;
		List<string> vecLine;
		vecLine = HS_ByteRead.readCsvLine( strContent, ref contentOffset );
		if(vecLine.Count != $ColCount$)
		{
			Debug.Log("$Template$.json 中列数量与生成的代码不匹配!");
			return false;
		}
$CheckColName$
		while(true)
		{
			vecLine = HS_ByteRead.readCsvLine( strContent, ref contentOffset );
			if((int)vecLine.Count == 0 )
				break;
			if((int)vecLine.Count != (int)$ColCount$)
			{
				return false;
			}
			$Template$Element member = new $Template$Element();
$ReadCsvColValue$
			member.IsValidate = true;
			_VecAllElements.Add(member);
			_MapElements[member.$PrimaryKey$] = member;
		}
		return true;
	}

	public bool LoadJson(string strContent)
	{
	    JsonData jsonData = JsonMapper.ToObject(strContent);
	    for (int i = 0; i < jsonData.Count; ++i)
	    {
	    	JsonData jd = jsonData[i];
	    	if(jd.Keys.Count != $ColCount$)
            {
                Debug.Log("$Template$.json中列数量与生成的代码不匹配!");
                return false;
            }
            
	        $Template$Element member = new $Template$Element();
$ReadJsonColValue$

	        member.IsValidate = true;
            _VecAllElements.Add(member);
            _MapElements[member.$PrimaryKey$] = member;
	    }
	    return true;
	}
};
