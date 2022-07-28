/* eslint-disable no-var */
import {
  PowerBiWorkspace,
  PowerBiDashboard,
  PowerBiReport,
  PowerBiDataset,
  PowerBiDashboardTile
}
  from "./../models/PowerBiModels";

import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';

import * as powerbi from "powerbi-client";
import * as pbimodels from "powerbi-models";
import { IPowerBiElement } from 'service';

require('powerbi-models');
require('powerbi-client');

export class PowerBiService {

  private static powerbiApiResourceId = "https://analysis.windows.net/powerbi/api";

  private static workspacesUrl = "https://api.powerbi.com/v1.0/myorg/groups/";


 //private static adalAccessTokenStorageKey: string = "adal.access.token.keyhttps://analysis.windows.net/powerbi/api";
  private static accessToken2 ="eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSIsImtpZCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvMGQ2OWI5MWQtNTg0OS00NDg2LThkMWQtM2YwOTNmMzcwNTRkLyIsImlhdCI6MTY1ODkzMTc0NSwibmJmIjoxNjU4OTMxNzQ1LCJleHAiOjE2NTg5MzU3MTcsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJFMlpnWUJCOTA1VmxibEZVbjFnWS9YSUpXM1QyeWI4bmM1bDNNY1JXdk5tczl1emxWeDRBIiwiYW1yIjpbInB3ZCJdLCJhcHBpZCI6IjU1ZDU5ZTNkLWEwNDAtNGMyYS05ZTkyLWZhNGJkNTQxM2FlMCIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoia2hhbGRpIiwiZ2l2ZW5fbmFtZSI6ImhhY2hpbSIsImlwYWRkciI6IjU0Ljg2LjUwLjEzOSIsIm5hbWUiOiJoYWNoaW0ga2hhbGRpIiwib2lkIjoiYWJiMzdlOTYtNDZmMi00Y2RkLThlZjMtYjA5M2UwZjhjNWRlIiwicHVpZCI6IjEwMDMyMDAyMDQwOTM1ODQiLCJyaCI6IjAuQVlJQUhibHBEVWxZaGtTTkhUOEpQemNGVFFrQUFBQUFBQUFBd0FBQUFBQUFBQUNWQU9vLiIsInNjcCI6IkFwcC5SZWFkLkFsbCBDYXBhY2l0eS5SZWFkLkFsbCBDYXBhY2l0eS5SZWFkV3JpdGUuQWxsIENvbnRlbnQuQ3JlYXRlIERhc2hib2FyZC5SZWFkLkFsbCBEYXNoYm9hcmQuUmVhZFdyaXRlLkFsbCBEYXRhZmxvdy5SZWFkLkFsbCBEYXRhZmxvdy5SZWFkV3JpdGUuQWxsIERhdGFzZXQuUmVhZC5BbGwgRGF0YXNldC5SZWFkV3JpdGUuQWxsIEdhdGV3YXkuUmVhZC5BbGwgR2F0ZXdheS5SZWFkV3JpdGUuQWxsIFJlcG9ydC5SZWFkLkFsbCBSZXBvcnQuUmVhZFdyaXRlLkFsbCBTdG9yYWdlQWNjb3VudC5SZWFkLkFsbCBTdG9yYWdlQWNjb3VudC5SZWFkV3JpdGUuQWxsIFdvcmtzcGFjZS5SZWFkLkFsbCBXb3Jrc3BhY2UuUmVhZFdyaXRlLkFsbCIsInN1YiI6IkdORWs3dmFVaTROYTdUNXJncjBzU1dtWllVRVVoak5ua1lCNGtvbnAxclkiLCJ0aWQiOiIwZDY5YjkxZC01ODQ5LTQ0ODYtOGQxZC0zZjA5M2YzNzA1NGQiLCJ1bmlxdWVfbmFtZSI6ImhhY2hpbUB0eXRsay5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJoYWNoaW1AdHl0bGsub25taWNyb3NvZnQuY29tIiwidXRpIjoiUmhGTGtBWnRqVXlEY2MwMmN0ZzhBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il19.un2r0GHPWwHZ72EDV0OJesYK7Bogm_t4OUT4jQryMZZP9ajrF8uGczg6N9Y2j62fkfFY1F9Hix1W5XwVXtyHGL9ELhWnJzjVzOqT5MjiHEp698b0YT7kn6eKQ9rtiLeSwDdMlyXKGHbBnGD9HpmS3x-Z_iTJ5bGjS2a66kg9FBsNr8D8nwBeUI4eUdM7wrN8n27hBd-naF5UC8WvNtwCnzzxs_Gtwt5WEIgVFnPFgW-9mCYBcvXJK3csYnmtiSAMMTvtKi0aYPkoGnYe7fK_o3St2TkTGVhm93cMNOgrkEgVwUFBc0hsE9g3g0DSTZwX9mBUHg7OxKZ26gQG4l8BWA"
  // we have  a problem with the acces token that we need to get from session storage !
  // Should we change window.sessionStorage[PowerBiService.adalAccessTokenStorageKey] with another value ????
  public static GetWorkspaces = (serviceScope: ServiceScope): Promise<PowerBiWorkspace[]> => {
    let pbiClient: AadHttpClient = new AadHttpClient(serviceScope, PowerBiService.powerbiApiResourceId);
    
    var reqHeaders: HeadersInit = new Headers();
    reqHeaders.append("Accept", "*");
    return pbiClient.get(PowerBiService.workspacesUrl, AadHttpClient.configurations.v1, { headers: reqHeaders })
      .then((response: HttpClientResponse): Promise<any> => {
        return response.json();
      })
      .then((workspacesOdataResult: any): PowerBiWorkspace[] => {
        return workspacesOdataResult.value;
      });
  }

  public static GetReports = (serviceScope: ServiceScope, workspaceId: string): Promise<PowerBiReport[]> => {

    let reportsUrl = PowerBiService.workspacesUrl + workspaceId + "/reports/";

    let pbiClient: AadHttpClient = new AadHttpClient(serviceScope, PowerBiService.powerbiApiResourceId);
  
    var reqHeaders: HeadersInit = new Headers();
    reqHeaders.append("Accept", "*");
    return pbiClient.get(reportsUrl, AadHttpClient.configurations.v1, { headers: reqHeaders })
      .then((response: HttpClientResponse): Promise<any> => {
        return response.json();
      })
      .then((reportsOdataResult: any): PowerBiReport[] => {
        return reportsOdataResult.value.map((report: PowerBiReport) => {
          return {
            id: report.id,
            embedUrl: report.embedUrl,
            name: report.name,
            webUrl: report.webUrl,
            datasetId: report.datasetId,
            accessToken:PowerBiService.accessToken2
            //accessToken: window.sessionStorage[PowerBiService.adalAccessTokenStorageKey]
          };
        });
      });
  }


  public static GetReport = (serviceScope: ServiceScope, workspaceId: string, reportId: string): Promise<PowerBiReport> => {
    let reportUrl = PowerBiService.workspacesUrl + workspaceId + "/reports/" + reportId + "/";
    let pbiClient: AadHttpClient = new AadHttpClient(serviceScope, PowerBiService.powerbiApiResourceId);
    var reqHeaders: HeadersInit = new Headers();
    reqHeaders.append("Accept", "*");
    return pbiClient.get(reportUrl, AadHttpClient.configurations.v1, { headers: reqHeaders })
      .then((response: HttpClientResponse): Promise<any> => {
        return response.json();
      })
      .then((reportsOdataResult: any): PowerBiReport => {
        return {
          id: reportsOdataResult.id,
          embedUrl: reportsOdataResult.embedUrl,
          name: reportsOdataResult.name,
          webUrl: reportsOdataResult.webUrl,
          datasetId: reportsOdataResult.datasetId,
          accessToken:PowerBiService.accessToken2
          //accessToken: window.sessionStorage[PowerBiService.adalAccessTokenStorageKey]
        };
      });
  }

  public static GetDashboards = (serviceScope: ServiceScope, workspaceId: string): Promise<PowerBiDashboard[]> => {
    let dashboardsUrl = PowerBiService.workspacesUrl + workspaceId + "/dashboards/";
    let pbiClient: AadHttpClient = new AadHttpClient(serviceScope, PowerBiService.powerbiApiResourceId);
    var reqHeaders: HeadersInit = new Headers();
    reqHeaders.append("Accept", "*");
    return pbiClient.get(dashboardsUrl, AadHttpClient.configurations.v1, { headers: reqHeaders })
      .then((response: HttpClientResponse): Promise<any> => {
        return response.json();
      })
      .then((dashboardsOdataResult: any): PowerBiDashboard[] => {
        return dashboardsOdataResult.value.map((dashboard: PowerBiDashboard) => {
          return {
            id: dashboard.id,
            embedUrl: dashboard.embedUrl,
            displayName: dashboard.displayName,
            accessToken:PowerBiService.accessToken2
            //accessToken: window.sessionStorage[PowerBiService.adalAccessTokenStorageKey]
          };
        });
      });
  }
  
  public static GetDashboard = (serviceScope: ServiceScope, workspaceId: string, dashboardId: string): Promise<PowerBiDashboard> => {
    let dashboardUrl = PowerBiService.workspacesUrl + workspaceId + "/dashboards/" + dashboardId + "/";
    let pbiClient: AadHttpClient = new AadHttpClient(serviceScope, PowerBiService.powerbiApiResourceId);
    var reqHeaders: HeadersInit = new Headers();
    reqHeaders.append("Accept", "*");
    return pbiClient.get(dashboardUrl, AadHttpClient.configurations.v1, { headers: reqHeaders })
      .then((response: HttpClientResponse): Promise<any> => {
        return response.json();
      })
      .then((dashboardOdataResult: any): PowerBiDashboard => {
        return {
          id: dashboardOdataResult.id,
          embedUrl: dashboardOdataResult.embedUrl,
          displayName: dashboardOdataResult.displayName,
          accessToken:PowerBiService.accessToken2
          //accessToken: window.sessionStorage[PowerBiService.adalAccessTokenStorageKey]
        };
      });
  }

}