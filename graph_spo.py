"""
Microsoft Graph API SharePoint Doc
https://docs.microsoft.com/ja-jp/graph/api/resources/sharepoint?view=graph-rest-1.0
"""
import requests, json, re


class SpoAPI:
    """
    認証情報等
    """
    def __init__(self):
        self.__tenant_id = "XXXX"
        self.__app_id = "XXXX"
        self.__secret = "XXXX"
        self.__site_id = "XXXXX"

    """
    apiのアクセストークンを取得する間数(private)
    """
    def __get_token(self):
        auth_url = "https://login.microsoftonline.com/"+ self.__tenant_id +"/oauth2/v2.0/token"
        headers = {
            'Accept': 'application/json',
        }

        payload = {
            'client_id': self.__app_id,
            'scope': 'https://graph.microsoft.com/.default',
            'grant_type': 'client_credentials',
            'client_secret': self.__secret
        }
        res = requests.post(auth_url, headers=headers, data=payload)
        return res.json()['access_token']

    """
    ファイルを取得する間数
    """
    def access_graph(self):
        url = "https://graph.microsoft.com/v1.0/sites/"+ self.__site_id +"/drive/items/root/children"
        token = self.__get_token()
        headers = {
            'Authorization': 'Bearer %s' % token
        }
        res = requests.get(url, headers=headers)
        payload = []

        for i in res.json()['value']:
            item_id = "{" + re.findall("(?<=\=).+?(?=\&)", i['@microsoft.graph.downloadUrl'])[0] + "}"
            payload.append(
                    {
                        'name': i['name'],
                        'iframe_url': "https://4msvlx.sharepoint.com/sites/resttest/_layouts/15/Doc.aspx?sourcedoc="+ item_id +"&amp;action=embedview"
                    }
            )

        return payload



if __name__=='__main__':
    spo_file = SpoAPI().access_graph()
    context = {'spo_file': spo_file}
    print(json.dumps(context, indent=2, ensure_ascii=False))

"""
実行結果

{
  "spo_file": [
    {
      "name": "Share PointとmFINDアプリケーションの接続に関して.pptx",
      "iframe_url": "https://4msvlx.sharepoint.com/sites/resttest/_layouts/15/Doc.aspx?sourcedoc={ddb036b4-9c48-4751-86d4-a072cccacfdb}&amp;action=embedview"
    },
    {
      "name": "社内教育受講報告書(0821-0920)_社員番号_名前.doc",
      "iframe_url": "https://4msvlx.sharepoint.com/sites/resttest/_layouts/15/Doc.aspx?sourcedoc={8bc92639-e8ab-49c8-86f6-e8cf2b9408d1}&amp;action=embedview"
    }
  ]
}
"""
