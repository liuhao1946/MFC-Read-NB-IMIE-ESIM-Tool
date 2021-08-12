
#include "pch.h"
#include "app_common.h"


void app_get_cstring_unit(CString str, BYTE *p, int len)
{
	int i, len_str;

	len_str = str.GetLength();

	if (len_str > len)
	{
		len_str = len - 1;
	}

	for (i = 0; i < len_str; i++)
	{
		*p++ = str.GetAt(i);
	}

	*p = 0;
}

void app_get_exe_path(char *p, int len)
{
	CString file_path = _T("");
	TCHAR exePath[MAX_PATH]; // MAX_PATH
	int idx = 0, idx_pre = 0, i = 0, str_len;
	char *p_src;


	GetModuleFileName(NULL, exePath, MAX_PATH);//
	file_path = exePath;

	while (i < 1000)
	{
		i++;
		idx = file_path.Find('\\', idx);

		if (idx == -1)
		{
			break;
		}
		file_path.Insert(idx, '\\');
		idx += 2;
		idx_pre = idx;
	}

	file_path.SetAt(idx_pre, 0);

	app_get_cstring_unit(file_path, (BYTE *)p, len);
}

void app_get_cstring_unit(CString str, char *p, int len)
{
	int i, len_str;

	len_str = str.GetLength();

	if (len_str > len)
	{
		len_str = len - 1;
	}

	for (i = 0; i < len_str; i++)
	{
		*p++ = str.GetAt(i);
	}

	*p = 0;
}