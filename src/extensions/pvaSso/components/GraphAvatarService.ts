import MSALWrapper from './MSALWrapper';

export class GraphAvatarService {
  private msalWrapper: MSALWrapper;
  private graphEndpoint: string = 'https://graph.microsoft.com/v1.0';
  
  constructor(clientId: string, authority: string) {
    this.msalWrapper = new MSALWrapper(clientId, authority);
  }
  
  /**
   * Retrieves the user's avatar from Microsoft Graph
   * @param userEmail The email of the user
   * @returns A promise that resolves to a base64-encoded data URL of the user's avatar, or null if retrieval fails
   */
  public async getUserAvatar(userEmail: string): Promise<string | null> {
    try {
      // Get token for Microsoft Graph API with appropriate scope
      const token = await this.msalWrapper.acquireAccessToken(['https://graph.microsoft.com/User.Read'], userEmail);
      
      if (!token || !token.accessToken) {
        console.error('Failed to acquire token for Graph API');
        return null;
      }
      
      // Call Microsoft Graph API to get user photo
      const response = await fetch(`${this.graphEndpoint}/me/photo/$value`, {
        headers: {
          'Authorization': `Bearer ${token.accessToken}`
        }
      });
      
      if (!response.ok) {
        console.error(`Failed to retrieve user photo: ${response.status}`);
        return null;
      }
      
      // Convert photo to base64
      const photoArrayBuffer = await response.arrayBuffer();
      const photoBase64 = this.arrayBufferToBase64(photoArrayBuffer);
      
      return `data:image/jpeg;base64,${photoBase64}`;
    } catch (error) {
      console.error('Error retrieving user avatar:', error);
      return null;
    }
  }
  
  /**
   * Converts an ArrayBuffer to a base64 string
   * @param buffer The ArrayBuffer to convert
   * @returns A base64 string
   */
  private arrayBufferToBase64(buffer: ArrayBuffer): string {
    const binary = Array.from(new Uint8Array(buffer))
      .map(byte => String.fromCharCode(byte))
      .join('');
    return btoa(binary);
  }
  
  /**
   * Fallback method to get a default avatar based on user initials
   * @param displayName The display name of the user
   * @returns A string with the user's initials
   */
  public getUserInitials(displayName: string): string {
    if (!displayName) return '';
    
    const nameParts = displayName.split(' ');
    if (nameParts.length === 1) {
      return nameParts[0].charAt(0).toUpperCase();
    }
    
    return (nameParts[0].charAt(0) + nameParts[nameParts.length - 1].charAt(0)).toUpperCase();
  }
}