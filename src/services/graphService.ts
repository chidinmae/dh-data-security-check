/**
 * Microsoft Graph API Service
 */

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export const fetchGraph = async (endpoint: string, accessToken: string, options: RequestInit = {}) => {
  const headers = new Headers();
  const bearer = `Bearer ${accessToken}`;

  headers.append("Authorization", bearer);

  const config = {
    ...options,
    headers: {
      ...options.headers,
      "Authorization": bearer
    }
  };

  return fetch(`${GRAPH_BASE_URL}${endpoint}`, config);
};

export const getSites = async (accessToken: string) => {
  const response = await fetchGraph("/sites?search=*", accessToken);
  if (!response.ok) throw new Error("Failed to fetch sites");
  const data = await response.json();
  return data.value;
};

export const getDrives = async (siteId: string, accessToken: string) => {
  const response = await fetchGraph(`/sites/${siteId}/drives`, accessToken);
  if (!response.ok) throw new Error("Failed to fetch drives");
  const data = await response.json();
  return data.value;
};

export const getDriveItems = async (driveId: string, accessToken: string) => {
  // Focus on text-based files in the root folder, recursively (simplified for POC)
  const response = await fetchGraph(`/drives/${driveId}/root/children?$expand=permissions`, accessToken);
  if (!response.ok) throw new Error("Failed to fetch drive items");
  const data = await response.json();
  return data.value;
};

export const getFileContent = async (driveId: string, itemId: string, accessToken: string) => {
  const response = await fetchGraph(`/drives/${driveId}/items/${itemId}/content`, accessToken);
  if (!response.ok) return null;
  return response.text();
};
