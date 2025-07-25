/*
 * Copyright 2025 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */

import core from '@actions/core';
import path from 'path';

// Base URL for DA Admin API – default, but will be overwritten in run()
let DA_API_BASE = 'https://admin.da.live';

async function daFetch(token, endpoint, options = {}) {
  const url = `${DA_API_BASE}${endpoint}`;
  core.info(`Fetching DA API endpoint: ${url}`);
  const defaultOptions = {
    headers: {
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      'Content-Type': 'application/json',
      Accept: 'application/json',
    },
  };

  const mergedOptions = {
    ...defaultOptions,
    ...options,
    headers: {
      ...defaultOptions.headers,
      ...options.headers,
    },
  };

  const res = await fetch(url, mergedOptions);

  if (!res.ok) {
    const errorText = await res.text();
    core.warning(`DA API error ${res.status}: ${errorText}`);
    throw new Error(`DA API error ${res.status}: ${errorText}`);
  }

  // Handle different response types
  const contentType = res.headers.get('content-type');
  if (contentType && contentType.includes('application/json')) {
    return res.json();
  }
  return res.text();
}

/**
 * Create a site (folder) in DA under given organisation using the source API
 * @param {string} token - Bearer token for DA API
 * @param {string} org - Organization name
 * @param {string} sitePath - Site path
 * @param {string} owner - Optional owner email for permissions
 */
async function createDASite(token, org, sitePath, owner = null) {
  try {
    // Create the folder structure by creating a README file in the site
    const folderEndpoint = `/source/${org}/${sitePath}`;
    
    // First check if the folder already exists
    try {
      const doesExistResponse = await daFetch(token, folderEndpoint, { method: 'GET' });
      core.info(`DA site ${sitePath} already exists in ${org}`);
      return doesExistResponse;
    } catch (error) {
      // Folder doesn't exist, create it with README file
      core.info(`Creating DA site: ${sitePath} in ${org}`);
    }

    // Create the folder by sending a POST request to the folder endpoint
    const createResponse = await daFetch(token, folderEndpoint, {
      method: 'POST',
      headers: {
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
      },
    });
    core.info(`Successfully created DA site: ${sitePath}`);
    const folderUrl = `https://da.live/#/${org}/${sitePath}`;

    // Output folder URL for manual sharing
    core.setOutput('folder_url', folderUrl);
    core.setOutput('folder_web_url', folderUrl);
    core.info(`📂 Folder URL: ${folderUrl}`);

    // Output test folder path for reference
    core.setOutput('test_folder_path', sitePath);
    core.info(`📁 Test folder to create: ${sitePath}`);

    // TODO: Set permission for owner
    if (owner) {
      // await setDAPermissions(token, org, siteName, owner);
    }

    return createResponse;

  } catch (error) {
    core.warning(`Failed to create DA site: ${error.message}`);
    core.setOutput('error_message', `❌ Error: Failed to create DA site: ${error.message}`);
    throw error;
  }
}

/**
 * Main function to create DA site and folder
 */
export async function run() {
  // the fallback on the env variables is for local testing
  const daImsToken       = core.getInput('da_ims_token') || process.env.DA_IMS_TOKEN; // Adobe IMS token
  const daAdminHost   = core.getInput('da_admin_host') || process.env.DA_ADMIN_HOST  || DA_API_BASE; // e.g., https://admin.da.live
  const daOrg         = core.getInput('da_org_path') || process.env.DA_ORG_PATH; // DA organisation name
  const daSitePath    = core.getInput('da_site_path') || process.env.DA_SITE_PATH; // site folder path, default='aemy-sites'
  const testFolderPath = core.getInput('test_folder_path') || process.env.TEST_FOLDER_PATH; // folder name to create for testing
  const daOwner       = core.getInput('da_owner') || process.env.DA_OWNER; // optional owner email for setting permissions

  const siteName = path.join(daSitePath, testFolderPath);

  // Update global base URL for DA API calls
  DA_API_BASE = daAdminHost.replace(/\/$/, ''); // remove trailing slash if present
  core.debug(`Using DA admin host: ${DA_API_BASE}`);

  try {
    // Use the provided Adobe IMS token
    const token = daImsToken;
    core.setOutput('access_token', token);

    // Create the DA site
    core.debug(`Creating DA site for org: ${daOrg}, siteName: ${testFolderPath}`);
    await createDASite(token, daOrg, siteName, daOwner);
    
    core.info(`✅ DA Site URL: https://da.live/#/${daOrg}/${siteName}`);
    core.debug(`DA site created: ${siteName} ${daOwner ? `with owner: ${daOwner}` : 'without owner'}`);

  } catch (error) {
    core.setFailed(`DA site creation failed: ${error.message}`);
    core.setOutput('error_message', `❌ Error: ${error.message}`);
    throw error;
  }
}

await run(); 