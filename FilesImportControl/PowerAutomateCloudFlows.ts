/**
 * PowerAutomateCloudFlows.ts
 * 
 * A standalone class for file operations using Power Automate Cloud Flows
 * instead of direct Azure SDK integration. This class provides upload, list,
 * and delete operations through Cloud Flow invocations.
 */

/**
 * Configuration interface for Power Automate Cloud Flows
 */
export interface ICloudFlowConfig {
  /** Base URL for the Power Automate environment */
  environmentUrl: string;
  /** Authentication token (Bearer token or connection string) */
  authToken: string;
  /** Flow ID for file upload operations */
  uploadFlowId: string;
  /** Flow ID for listing files operations */
  listFilesFlowId: string;
  /** Flow ID for file deletion operations */
  deleteFlowId: string;
  /** Flow ID for generating SAS URLs to view files */
  generateViewUrlFlowId: string;
  /** Container/folder path for file operations */
  containerPath: string;
  /** Optional timeout for flow execution (in milliseconds) */
  timeout?: number;
}

/**
 * Event object containing transfer progress information
 */
export interface TransferProgressEvent {
  /** Number of bytes loaded/transferred so far */
  loadedBytes: number;
  /** Total file size in bytes */
  totalBytes?: number;
  /** Progress percentage (0-100) */
  percentage?: number;
}

/**
 * Result object returned after file upload operation
 */
export interface IUploadResult {
  /** Name of the uploaded file */
  fileName: string;
  /** Full URL where the file can be accessed */
  url: string;
  /** Whether the upload was successful */
  success: boolean;
  /** Error message if upload failed */
  error?: string;
  /** Flow execution ID for tracking */
  flowRunId?: string;
}

/**
 * Information about an existing file in storage
 */
export interface IExistingFile {
  /** File name */
  name: string;
  /** File size in bytes */
  size: number;
  /** Full URL where the file can be accessed */
  url: string;
  /** Date when the file was last modified */
  lastModified: Date;
  /** Additional metadata returned by the flow */
  metadata?: Record<string, any>;
}

/**
 * Result object returned after generating view URL operation
 */
export interface IGenerateViewUrlResult {
  /** Name of the file for which the view URL was generated */
  fileName: string;
  /** SAS URL for viewing the file */
  viewUrl: string;
  /** Whether the operation was successful */
  success: boolean;
  /** Error message if operation failed */
  error?: string;
  /** Flow execution ID for tracking */
  flowRunId?: string;
}

/**
 * Result object returned after file deletion operation
 */
export interface IDeleteResult {
  /** Name of the file that was deleted */
  fileName: string;
  /** Whether the deletion was successful */
  success: boolean;
  /** Error message if deletion failed */
  error?: string;
  /** Flow execution ID for tracking */
  flowRunId?: string;
}

/**
 * Request payload for upload flow
 */
interface IUploadFlowRequest {
  fileName: string;
  fileContent: string; // Base64 encoded
  containerPath: string;
  folderName?: string;
  fileSize: number;
  contentType: string;
}

/**
 * Response from upload flow
 */
interface IUploadFlowResponse {
  success: boolean;
  fileName: string;
  url: string;
  error?: string;
  flowRunId: string;
}

/**
 * Request payload for list files flow
 */
interface IListFilesFlowRequest {
  containerPath: string;
  folderName?: string;
}

/**
 * Response from list files flow
 */
interface IListFilesFlowResponse {
  success: boolean;
  files: Array<{
    name: string;
    size: number;
    url: string;
    lastModified: string;
    metadata?: Record<string, any>;
  }>;
  error?: string;
  flowRunId: string;
}

/**
 * Request payload for generate view URL flow
 */
interface IGenerateViewUrlFlowRequest {
  containerPath: string;
  filePath: string;
}

/**
 * Response from generate view URL flow
 */
interface IGenerateViewUrlFlowResponse {
  success: boolean;
  fileName: string;
  viewUrl: string;
  error?: string;
  flowRunId: string;
}

/**
 * Request payload for delete flow
 */
interface IDeleteFlowRequest {
  fileName: string;
  containerPath: string;
  folderName?: string;
}

/**
 * Response from delete flow
 */
interface IDeleteFlowResponse {
  success: boolean;
  fileName: string;
  error?: string;
  flowRunId: string;
}

/**
 * Utility function to convert File to Base64
 * @param file - File object to convert
 * @returns Promise resolving to Base64 string
 */
const fileToBase64 = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (reader.result) {
        // Remove the data URL prefix (e.g., "data:image/png;base64,")
        const base64 = (reader.result as string).split(',')[1];
        resolve(base64);
      } else {
        reject(new Error('Failed to read file'));
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
};

/**
 * Utility function to invoke a Power Automate Cloud Flow
 * @param flowUrl - Complete URL to the flow trigger endpoint
 * @param payload - Request payload to send to the flow
 * @param config - Flow configuration containing auth details
 * @returns Promise resolving to flow response
 */
const invokeCloudFlow = async <TRequest, TResponse>(
  flowUrl: string,
  payload: TRequest,
  config: ICloudFlowConfig
): Promise<TResponse> => {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => {
    controller.abort();
  }, config.timeout || 100000); // Default 100 seconds (Power Automate timeout limit)

  try {
    const response = await fetch(flowUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${config.authToken}`,
      },
      body: JSON.stringify(payload),
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      throw new Error(`Flow execution failed: ${response.status} ${response.statusText}`);
    }

    const result = await response.json();
    return result as TResponse;
  } catch (error) {
    clearTimeout(timeoutId);
    if (error instanceof Error) {
      if (error.name === 'AbortError') {
        throw new Error('Flow execution timed out');
      }
      throw error;
    }
    throw new Error('Unknown error occurred during flow execution');
  }
};

/**
 * Constructs the complete flow trigger URL
 * @param config - Flow configuration
 * @param flowId - ID of the specific flow to invoke
 * @returns Complete flow trigger URL
 */
const buildFlowUrl = (config: ICloudFlowConfig, flowId: string): string => {
  const baseUrl = config.environmentUrl.endsWith('/')
    ? config.environmentUrl.slice(0, -1)
    : config.environmentUrl;
  return `${baseUrl}/_api/cloudflow/v1.0/trigger/${flowId}`;
};

/**
 * Uploads a single file using Power Automate Cloud Flow
 * @param file - File to upload (null if no file selected)
 * @param config - Cloud Flow configuration
 * @param folderName - Optional folder name where the file should be uploaded
 * @param onProgress - Optional callback function to track upload progress
 * @returns Promise resolving to upload result or null if no file provided
 */
export const uploadFileToCloudFlow = async (
  file: File | null,
  config: ICloudFlowConfig,
  folderName?: string,
  onProgress?: (progress: TransferProgressEvent) => void
): Promise<IUploadResult | null> => {
  if (!file) return null;

  try {
    // Simulate progress during file encoding
    if (onProgress) {
      onProgress({
        loadedBytes: 0,
        totalBytes: file.size,
        percentage: 0,
      });
    }

    // Convert file to Base64
    const base64Content = await fileToBase64(file);

    // Simulate progress after encoding
    if (onProgress) {
      onProgress({
        loadedBytes: file.size * 0.5,
        totalBytes: file.size,
        percentage: 50,
      });
    }

    // Prepare request payload
    const uploadRequest: IUploadFlowRequest = {
      fileName: file.name,
      fileContent: base64Content,
      containerPath: config.containerPath,
      folderName: folderName,
      fileSize: file.size,
      contentType: file.type || 'application/octet-stream',
    };

    // Build flow URL
    const flowUrl = buildFlowUrl(config, config.uploadFlowId);

    // Invoke the upload flow
    const response = await invokeCloudFlow<IUploadFlowRequest, IUploadFlowResponse>(
      flowUrl,
      uploadRequest,
      config
    );

    // Simulate progress completion
    if (onProgress) {
      onProgress({
        loadedBytes: file.size,
        totalBytes: file.size,
        percentage: 100,
      });
    }

    // Return standardized result
    return {
      fileName: response.fileName,
      url: response.url,
      success: response.success,
      error: response.error,
      flowRunId: response.flowRunId,
    };
  } catch (error) {
    console.error('Cloud Flow upload error:', error);
    return {
      fileName: file.name,
      url: "",
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
};

/**
 * Uploads multiple files using Power Automate Cloud Flow sequentially
 * @param files - Array of files to upload
 * @param config - Cloud Flow configuration
 * @param folderName - Optional folder name where the files should be uploaded
 * @param onProgress - Optional callback for tracking individual file upload progress
 * @returns Promise resolving to array of upload results for each file
 */
export const uploadMultipleFilesToCloudFlow = async (
  files: File[],
  config: ICloudFlowConfig,
  folderName?: string,
  onProgress?: (fileName: string, progress: TransferProgressEvent) => void
): Promise<IUploadResult[]> => {
  const results: IUploadResult[] = [];

  // Process files sequentially to avoid overwhelming the flow
  for (const file of files) {
    const result = await uploadFileToCloudFlow(
      file,
      config,
      folderName,
      // Wrap progress callback to include filename for multi-file tracking
      onProgress ? (progress: TransferProgressEvent) => onProgress(file.name, progress) : undefined
    );
    if (result) {
      results.push(result);
    }
  }

  return results;
};

/**
 * Lists all existing files using Power Automate Cloud Flow
 * @param config - Cloud Flow configuration
 * @param folderName - Optional folder name to filter results
 * @returns Promise resolving to array of existing file information
 */
export const listExistingFilesFromCloudFlow = async (
  config: ICloudFlowConfig,
  folderName?: string
): Promise<IExistingFile[]> => {
  try {
    // Prepare request payload
    const listRequest: IListFilesFlowRequest = {
      containerPath: config.containerPath,
      folderName: folderName,
    };

    // Build flow URL
    const flowUrl = buildFlowUrl(config, config.listFilesFlowId);

    // Invoke the list files flow
    const response = await invokeCloudFlow<IListFilesFlowRequest, IListFilesFlowResponse>(
      flowUrl,
      listRequest,
      config
    );

    if (!response.success) {
      console.error('List files flow failed:', response.error);
      return [];
    }

    // Convert response to standardized format
    return response.files.map(file => ({
      name: file.name,
      size: file.size,
      url: file.url,
      lastModified: new Date(file.lastModified),
      metadata: file.metadata,
    }));
  } catch (error) {
    console.error('Error listing files from Cloud Flow:', error);
    return [];
  }
};

/**
 * Generates a SAS URL for viewing a specific file using Power Automate Cloud Flow
 * @param fileName - Name of the file to generate view URL for
 * @param config - Cloud Flow configuration
 * @param folderName - Optional folder name where the file is located
 * @returns Promise resolving to view URL generation result
 */
export const generateViewUrlFromCloudFlow = async (
  fileName: string,
  config: ICloudFlowConfig,
  folderName?: string
): Promise<IGenerateViewUrlResult> => {
  try {
    // Prepare the file path based on folder structure
    const filePath = folderName ? `${folderName}/${fileName}` : fileName;
    
    // Prepare request payload
    const generateViewUrlRequest: IGenerateViewUrlFlowRequest = {
      containerPath: config.containerPath,
      filePath: filePath,
    };

    // Build flow URL
    const flowUrl = buildFlowUrl(config, config.generateViewUrlFlowId);

    // Invoke the generate view URL flow
    const response = await invokeCloudFlow<IGenerateViewUrlFlowRequest, IGenerateViewUrlFlowResponse>(
      flowUrl,
      generateViewUrlRequest,
      config
    );

    return {
      fileName: response.fileName,
      viewUrl: response.viewUrl,
      success: response.success,
      error: response.error,
      flowRunId: response.flowRunId,
    };
  } catch (error) {
    console.error(`Error generating view URL for file ${fileName} from Cloud Flow:`, error);
    return {
      fileName,
      viewUrl: "",
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
};

/**
 * Deletes a specific file using Power Automate Cloud Flow
 * @param fileName - Name of the file to delete
 * @param config - Cloud Flow configuration
 * @param folderName - Optional folder name where the file is located
 * @returns Promise resolving to deletion result
 */
export const deleteFileFromCloudFlow = async (
  fileName: string,
  config: ICloudFlowConfig,
  folderName?: string
): Promise<IDeleteResult> => {
  try {
    // Prepare request payload
    const deleteRequest: IDeleteFlowRequest = {
      fileName: fileName,
      containerPath: config.containerPath,
      folderName: folderName,
    };

    // Build flow URL
    const flowUrl = buildFlowUrl(config, config.deleteFlowId);

    // Invoke the delete flow
    const response = await invokeCloudFlow<IDeleteFlowRequest, IDeleteFlowResponse>(
      flowUrl,
      deleteRequest,
      config
    );

    return {
      fileName: response.fileName,
      success: response.success,
      error: response.error,
      flowRunId: response.flowRunId,
    };
  } catch (error) {
    console.error(`Error deleting file ${fileName} from Cloud Flow:`, error);
    return {
      fileName,
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
};

/**
 * Utility function to validate Cloud Flow configuration
 * @param config - Configuration to validate
 * @returns Boolean indicating if configuration is valid
 */
export const validateCloudFlowConfig = (config: ICloudFlowConfig): boolean => {
  return !!(
    config.environmentUrl &&
    config.authToken &&
    config.uploadFlowId &&
    config.listFilesFlowId &&
    config.deleteFlowId &&
    config.generateViewUrlFlowId &&
    config.containerPath
  );
};

// Export the main upload function as default for backward compatibility
export default uploadFileToCloudFlow;
