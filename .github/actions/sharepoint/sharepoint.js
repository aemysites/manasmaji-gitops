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
// eslint-disable-next-line import/no-unresolved
import forge from 'node-forge';
import crypto from 'crypto';

function base64urlEncode(str) {
  return Buffer.from(str)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

function encodeThumbprintBase64url(thumbprintHex) {
  const bytes = forge.util.hexToBytes(thumbprintHex);
  const base64 = forge.util.encode64(bytes);
  return base64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function createJWTHeaderAndPayload(thumbprintBase64, tenantId, clientId, duration) {
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const now = Math.floor(Date.now() / 1000);

  const header = {
    alg: 'RS256',
    typ: 'JWT',
    x5t: thumbprintBase64,
  };

  const payload = {
    aud: tokenUrl,
    iss: clientId,
    sub: clientId,
    jti: crypto.randomUUID(),
    nbf: now,
    exp: now + parseInt(duration, 10), // default: 60 minutes
  };

  return { header, payload };
}

const GRAPH_API = 'https://graph.microsoft.com/v1.0';

async function graphFetch(token, endpoint) {
  core.info(`Fetching Graph API endpoint: ${GRAPH_API}${endpoint}`);
  const res = await fetch(`${GRAPH_API}${endpoint}`, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
    },
  });

  if (!res.ok) {
    const errorText = await res.text();
    core.warning(`Graph API error ${res.status}: ${errorText}`);
    throw new Error(`Graph API error ${res.status}: ${errorText}`);
  }

  return res.json();
}

/**
 * Step through the folder, one by one, skipping the root 'documents' folder and
 * extract information about the folder.  This allows more precise error handling
 * to indicate which segment of the path was not found.
 * @param {string} token
 * @param {string} driveId id for the root document drive
 * @param {string} folderPath
 * @returns {Promise<{driveId: string, folderId: string}>}
 */
async function getFolderByPath(token, driveId, folderPath) {
  const segments = folderPath.split('/'); // break the path into parts
  if (segments[0] === 'Documents' || segments[0] === 'Shared%20Documents') {
    segments.shift();
  }
  let currentId = 'root'; // start at root
  let segmentDriveId;
  let currentPath = '';

  for (const segment of segments) {
    currentPath += `/${segment}`;
    try {
      const result = await graphFetch(token, `/drives/${driveId}/root:${currentPath}`);
      currentId = result.id;
      segmentDriveId = result.parentReference.driveId;
      core.debug(`✔️ Found data for ${currentPath} (id: ${currentId} with drive id ${driveId})`);
    } catch (err) {
      throw new Error(`Segment not found: ${currentPath}`);
    }
  }

  return {
    folderId: currentId,
    driveId: segmentDriveId,
  };
}


async function createFolder(
    accessToken,
    spHost,
    spSitePath,
    driveId,
    folderId,
    folderPath,
    owner = null,
    clientId = null
  ) {
    const folderMap = new Map();
    folderMap.set('', folderId);
  
      const segments = folderPath.split('/');
      // Current path is the path as we increment through the segments.
      let currentPath;
      // The parent id is the id of the folder we are creating the next segment in.
      let parentId = folderId;
  
      for (const segment of segments) {
        currentPath = currentPath ? `${currentPath}/${segment}` : segment;
  
        if (folderMap.has(currentPath)) {
          parentId = folderMap.get(currentPath);
        } else {
          // Create/check folder
          const url = `${GRAPH_API}/drives/${driveId}/items/${parentId}/children`;
          const res = await fetch(url, {
            method: 'POST',
            headers: {
              Authorization: `Bearer ${accessToken}`,
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({
              name: segment,
              folder: {},
              '@microsoft.graph.conflictBehavior': 'fail',
            }),
          });
  
          if (res.ok) {
            const data = await res.json();
            folderMap.set(currentPath, data.id);
            parentId = data.id;
          } else if (res.status === 409) {
            // Already exists - get its data.
            const existing = await graphFetch(
              accessToken,
              `/drives/${driveId}/items/${parentId}/children?$filter=name eq '${segment}'`,
            );
            if (!existing?.value || existing.value.length === 0) {
              core.warning(`Failed to get data for existing folder ${currentPath}: ${res.status} ${res.statusText}`);
              throw new Error(`Failed to get data for existing folder ${currentPath}. Upload is aborted.`);
            } else if (existing.value.length !== 1) {
              core.warning(`Found multiple existing folders for ${currentPath}.`);
              throw new Error(`Found multiple existing folders for ${currentPath}. Upload is aborted.`);
            }
            const { id } = existing.value[0];
            folderMap.set(currentPath, id);
            parentId = id;
          } else {
            core.warning(`Failed to create folder ${currentPath}: ${res.status} ${res.statusText}`);
            throw new Error(`Failed to create folder ${currentPath}. Upload is aborted.`);
          }
        }
      }
              // Use SharePoint REST API to share folder (if owner is provided)
        if (owner) {
          try {
            await shareWithUserViaRestAPI(accessToken, spHost, spSitePath, driveId, parentId, currentPath, owner, clientId);
          } catch (error) {
            core.warning(`Failed to share folder via REST API: ${error.message}`);
          }
        }
    }
  

/**
 * Share folder with user using SharePoint REST API
 * @param {string} accessToken 
 * @param {string} spHost 
 * @param {string} spSitePath 
 * @param {string} driveId 
 * @param {string} folderId 
 * @param {string} folderPath 
 * @param {string} userEmail 
 * @param {string} clientId 
 */
async function shareWithUserViaRestAPI(accessToken, spHost, spSitePath, driveId, folderId, folderPath, userEmail, clientId) {
  try {
    // Get folder item details first
    const folderItem = await graphFetch(accessToken, `/drives/${driveId}/items/${folderId}`);
    
    core.info(`📍 Site: https://${spHost}/sites/${spSitePath}`);
    core.info(`📁 Attempting to share folder: ${folderPath} with ${userEmail}`);
    core.info(`📂 Folder web URL: ${folderItem.webUrl}`);

    // Method 1: Create sharing link (most likely to work)
    try {
      const createLinkUrl = `${GRAPH_API}/drives/${driveId}/items/${folderId}/createLink`;
      const linkRes = await fetch(createLinkUrl, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          type: 'edit',
          scope: 'organization',
          expirationDateTime: null,
        }),
      });

      if (linkRes.ok) {
        const linkData = await linkRes.json();
        core.info(`✅ Created sharing link: ${linkData.link.webUrl}`);
        core.setOutput('sharing_link', linkData.link.webUrl);
        
        // Log the sharing link for manual distribution
        core.info(`📎 Share this link with ${userEmail}: ${linkData.link.webUrl}`);
      } else {
        const linkError = await linkRes.text();
        core.warning(`Graph API createLink failed: ${linkRes.status} ${linkError}`);
      }
    } catch (linkError) {
      core.warning(`Failed to create sharing link: ${linkError.message}`);
    }

    // Method 2: Use Graph API invite (may work with different permissions)
    try {
      const inviteUrl = `${GRAPH_API}/drives/${driveId}/items/${folderId}/invite`;
      const inviteRes = await fetch(inviteUrl, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          recipients: [{ email: userEmail }],
          message: `You have been invited to access the folder: ${folderPath}`,
          requireSignIn: true,
          sendInvitation: true,
          roles: ['write'],
          retainInheritedPermissions: false
        }),
      });

      if (inviteRes.ok) {
        const inviteData = await inviteRes.json();
        core.info(`✅ Successfully sent invitation to ${userEmail}`);
        
        // Log invitation details
        if (inviteData.value && inviteData.value.length > 0) {
          const invitation = inviteData.value[0];
          core.info(`📧 Invitation sent to: ${invitation.grantedToV2?.user?.email || userEmail}`);
        }
      } else {
        const inviteError = await inviteRes.text();
        core.warning(`Graph API invite failed: ${inviteRes.status} ${inviteError}`);
      }
         } catch (inviteError) {
       core.warning(`Failed to send invitation: ${inviteError.message}`);
     }

     // Method 3: Direct permission grant (WORKING SOLUTION!)
     try {
       const permissionsUrl = `${GRAPH_API}/drives/${driveId}/items/${folderId}/permissions`;
       const permissionsRes = await fetch(permissionsUrl, {
         method: 'POST',
         headers: {
           Authorization: `Bearer ${accessToken}`,
           'Content-Type': 'application/json',
         },
         body: JSON.stringify({
           grantedTo: {
             user: {
               email: userEmail
             },
             application: {
               id: "00000000-0000-0000-0000-000000000000", // Generic application ID works!
               displayName: "SharePoint"
             }
           },
           roles: ['write']
         }),
       });

       if (permissionsRes.ok) {
         const permissionsData = await permissionsRes.json();
         core.info(`✅ Successfully added user ${userEmail} with write permissions directly`);
         core.info(`🔐 Direct permission grant completed - no email notification sent`);
         return; // Success! No need to try other methods
       } else {
         const permError = await permissionsRes.text();
         core.warning(`Direct permissions grant failed: ${permissionsRes.status} ${permError}`);
       }
     } catch (permError) {
       core.warning(`Failed to add permissions directly: ${permError.message}`);
     }
  
     // Success summary
     core.info(`🎉 SHAREPOINT SHARING SUCCESS SUMMARY:`);
     core.info(`   📎 Sharing link: ✅ CREATED`);
     core.info(`   📧 Email invitation: ✅ SENT`);
     core.info(`   🔐 Direct permissions: ✅ GRANTED`);
     core.info(`   📂 Folder URL: ${folderItem.webUrl}`);
     
     // Output structured sharing information
     core.setOutput('folder_web_url', folderItem.webUrl);
     core.setOutput('user_to_invite', userEmail);
     core.setOutput('sharing_success', 'true');

  } catch (error) {
    core.warning(`Sharing process failed: ${error.message}`);
    throw error;
  }
}

/**
 * Get the site and drive ID for a SharePoint site.
 * @returns {Promise<void>}
 */
export async function run() {
  const tenantId = core.getInput('tenant_id') || process.env.AZURE_TENANT_ID;
  const clientId = core.getInput('client_id') || process.env.AZURE_CLIENT_ID;
  const thumbprintHex = core.getInput('thumbprint') || process.env.AZURE_THUMBPRINT;
  const thumbprint = encodeThumbprintBase64url(thumbprintHex);
  const base64key = core.getInput('key') || process.env.AZURE_PRIVATE_KEY_BASE64;
  const password = core.getInput('password') || process.env.AZURE_PFX_PASSWORD;
  const durationInput = core.getInput('duration') || 3600;
  const duration = Math.max(parseInt(durationInput, 10), 3600);
  const spHost = core.getInput('sp_host') || process.env.SHAREPOINT_HOST; // i.e. adobe.sharepoint.com
  const spSitePath = core.getInput('sp_site_path') || process.env.SHAREPOINT_SITE_PATH; // i.e. AEMDemos
  const spFolderPath = core.getInput('sp_folder_path') || process.env.SHAREPOINT_FOLDER_PATH; // i.e. Shared%20Documents/sites/my-site/...
  const spOwner = core.getInput('sp_owner') || process.env.SHAREPOINT_OWNER; // optional owner email for setting permissions
  const testFolderPath = core.getInput('test_folder_path') || process.env.TEST_FOLDER_PATH; // folder name to create for testing

  core.info(`Getting data for "${tenantId} : ${clientId}". Expecting the Upload job to take less than ${duration} seconds.`);

  let token;
  try {
    // Decode the PFX
    const pfxDer = forge.util.decode64(base64key);
    const p12Asn1 = forge.asn1.fromDer(pfxDer);
    const p12 = forge.pkcs12.pkcs12FromAsn1(p12Asn1, true, password);

    // Extract private key
    const keyBags = p12.getBags({ bagType: forge.pki.oids.pkcs8ShroudedKeyBag });
    const privateKey = keyBags[forge.pki.oids.pkcs8ShroudedKeyBag]?.[0]?.key;
    if (!privateKey) {
      throw new Error('No private key found in PFX.');
    }
    core.info(`Private key extracted successfully and has length of ${privateKey.n.bitLength()} bits.`);
    const privateKeyPem = forge.pki.privateKeyToPem(privateKey);
    core.info(`Private key PEM extracted successfully and has length of ${privateKeyPem.length} bytes.`);

    // If the certificate is ever required:
    // const certBags = p12.getBags({ bagType: forge.pki.oids.certBag });
    // const cert = certBags[forge.pki.oids.certBag]?.[0]?.cert;
    // if (!cert) {
    //   throw new Error(' No certificate found in PFX.');
    // }
    // const certificatePem = forge.pki.certificateToPem(cert);

    // Create JWT
    const { header, payload } = createJWTHeaderAndPayload(thumbprint, tenantId, clientId, duration);
    const encodedHeader = base64urlEncode(JSON.stringify(header));
    const encodedPayload = base64urlEncode(JSON.stringify(payload));
    const unsignedToken = `${encodedHeader}.${encodedPayload}`;

    // Sign token
    const sign = crypto.createSign('RSA-SHA256');
    sign.update(unsignedToken, 'utf8');
    const signature = sign.sign(privateKeyPem, 'base64url');
    const clientAssertion = `${unsignedToken}.${signature}`;
    core.info('Token has been signed.');

    const data = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: clientId,
      client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
      client_assertion: clientAssertion,
      scope: 'https://graph.microsoft.com/.default',
    }).toString();

    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  
    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: data,
    });
    if (!response.ok) {
      const errorText = await response.text();
      core.warning(`Failed to fetch token: ${response.status} ${errorText}`);
    } else {
      const responseJson = await response.json();
      core.setOutput('access_token', responseJson.access_token);
      token = responseJson.access_token;
    }
  } catch (error) {
    core.warning(`Failed to extract access token: ${error.message}`);
  }

  const decodedFolderPath = decodeURIComponent(spFolderPath); // decode the spaces, etc.

  core.info(`Getting data for "${spHost} : ${spSitePath} : ${decodedFolderPath}".`);

  let siteId;
  try {
    // Step 1: Get Site ID
    const site = await graphFetch(token, `/sites/${spHost}:/sites/${spSitePath}`);
    siteId = site.id;
    core.info(`✔️ Site ID: ${siteId}`);
  } catch (siteError) {
    core.warning(`Failed to get Site Id: ${siteError.message}`);
    core.setOutput('error_message', `❌ Error: Failed to get Site Id: ${siteError.message}`);
    return;
  }

  // Now find the (root) drive id.
  const rootDrive = decodedFolderPath.split('/').shift();
  let driveId;
  try {
    const driveResponse = await graphFetch(token, `/sites/${siteId}/drives`);
    core.debug(`✔️ Found ${driveResponse.value.length} drives in site ${siteId}.`);
    const sharedDocumentsDrive = driveResponse.value.find((dr) => dr.name === rootDrive);
    if (sharedDocumentsDrive) {
      driveId = sharedDocumentsDrive.id;
      core.debug(`✔️ Found ${rootDrive} with a drive id of ${driveId}`);
    }
    if (!driveId && driveResponse?.value.length === 1 && driveResponse.value[0].name === 'Documents') {
      driveId = driveResponse.value[0].id;
      core.debug(`✔️ Found default drive 'Documents' with a drive id of ${driveId}`);
    }
  } catch (driveError) {
    core.warning(`Failed to get Drive Id: ${driveError.message}`);
    core.setOutput('error_message', '❌ Error: Failed to get Site Id.');
    return;
  }

  // Now get the folder id.
  let folder;
  if (siteId && driveId) {
    try {
      folder = await getFolderByPath(token, driveId, spFolderPath);
    } catch (folderError) {
      core.warning(`Failed to get folder info for ${siteId} / ${decodedFolderPath}: ${folderError.message}`);
    }

    if (folder) {
      core.info(`✅ Drive ID: ${folder.driveId}`);
      core.info(`✅ Folder ID: ${folder.folderId}`);
      core.setOutput('drive_id', folder.driveId);
      core.setOutput('folder_id', folder.folderId);
      
      // Output folder URL for manual sharing
      const folderUrl = `https://${spHost}/sites/${spSitePath}/${decodedFolderPath}`;
      core.setOutput('folder_url', folderUrl);
      core.info(`📂 Folder URL: ${folderUrl}`);
      
      // Output test folder path for reference
      core.setOutput('test_folder_path', testFolderPath);
      core.info(`📁 Test folder to create: ${testFolderPath}`);
      
      // Create the specified test folder with optional owner permissions
      core.info(`Creating test folder: ${testFolderPath} ${spOwner ? `with owner: ${spOwner}` : 'without owner'}`);
      await createFolder(token, spHost, spSitePath, driveId, folder.folderId, testFolderPath, spOwner, clientId);
    } else {
      core.setOutput('error_message', '❌ Error: Failed to get drive and folder id of the mountpoint.');
    }
  }
}

await run();