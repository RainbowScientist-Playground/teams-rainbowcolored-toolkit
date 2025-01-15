// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Project } from "../utils/constants";
import { Env } from "../utils/env";
import {
  AppStudioCleanHelper,
  filterResourceGroupByName,
  deleteResourceGroupByName,
  GraphApiCleanHelper,
  SharePointApiCleanHelper,
  DevTunnelCleanHelper,
  M365TitleCleanHelper,
} from "../utils/cleanHelper";
import { getAppNamePrefix } from "../utils/nameUtil";
import { delay } from "../utils/retryHandler";

const appStudioAppNamePrefixList: string[] = [Project.namePrefix, "vs"];
const appNamePrefixList: string[] = [Project.namePrefix, "vs"];
const aadNamePrefixList: string[] = [Project.namePrefix, "vs"];
const rgNamePrefixList: string[] = [Project.namePrefix, "vs"];
const excludePrefix: string = getAppNamePrefix();

async function main() {
  const cleanService = await GraphApiCleanHelper.create(
    Env.cleanTenantId,
    Env.cleanClientId,
    Env.username,
    Env.password
  );

  console.log(`clean AAD (exclude ${excludePrefix})`);
  const aadList = await cleanService.listAad();
  if (aadList) {
    for (const aad of aadList) {
      if (!aad.displayName?.startsWith("delete")) {
        console.log(aad.displayName);
        await cleanService.deleteAad(aad.id!);
      }
    }
  }

  console.log(`clean teams app (exclude ${excludePrefix})`);
  const teamsUserId = await cleanService.getUserIdByName(Env.username);
  const teamsAppList = await cleanService.listTeamsApp(teamsUserId);
  if (teamsAppList) {
    for (const app of teamsAppList) {
      console.log(app?.teamsAppDefinition?.displayName);
      try {
        await cleanService.uninstallTeamsApp(teamsUserId, app?.id ?? "");
      } catch {
        console.log(
          `Failed to uninstall Teams App ${app?.teamsAppDefinition?.displayName}`
        );
      }
    }
  }

  console.log(`clean app in app studio`);
  const addStudioCleanService = await AppStudioCleanHelper.create(
    Env.cleanTenantId,
    Env.cleanClientId,
    Env.username,
    Env.password
  );
  const appStudioAppList = await addStudioCleanService.getAppsInAppStudio();
  if (appStudioAppList) {
    for (const app of appStudioAppList) {
      console.log(app?.displayName);
      try {
        await addStudioCleanService.deleteAppInAppStudio(app?.appDefinitionId);
      } catch {
        console.log(
          `Failed to delete Teams App ${app?.displayName} in App Studio`
        );
      }
    }
  }

  let retry: boolean;
  let count = 10;
  const total = count + 1;
  do {
    retry = false;
    console.log(`Start to try ${total - count} times`);
    const m365TitleCleanService = await M365TitleCleanHelper.create(
      Env.cleanTenantId,
      "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
      Env.username,
      Env.password
    );
    console.log(`clean M365 Titles (exclude ${excludePrefix})`);
    try {
      const acquisitions = await m365TitleCleanService.listAcquisitions();
      if (acquisitions) {
        for (const acquisition of acquisitions) {
          console.log(acquisition.titleDefinition.name);
          console.log(acquisition.titleId);
          const result = await m365TitleCleanService.unacquire(
            acquisition.titleId
          );
          if (!retry && result) {
            retry = true;
          }
        }
      }
    } catch (e: any) {
      console.log(`Get error: ${e.message}`);
      retry = true;
      if (count > 1) {
        // Retry after a short time if getting "Rate limit is exceeded"
        await delay(30 * 1000);
      }
    }

    count--;
  } while (retry && count > 0);
}

main()
  .then((_) => {
    console.log("Clean Job Done.");
  })
  .catch((error) => {
    console.error(error);
    process.exit(-1);
  });
