import * as React from "react";
import { useState, createRef } from "react";
import { Caption1, Button, CompoundButton, Spinner, FluentProvider, Theme, webLightTheme, ProgressBar, Text, Body1, Caption2, Skeleton, SkeletonItem } from "@fluentui/react-components";
import { CheckmarkFilled, DismissRegular, CheckmarkCircleFilled, ErrorCircleFilled, DeleteRegular, EyeRegular, CheckmarkRegular } from "@fluentui/react-icons";
import { getIcon } from "./iconsMapping";
import { ButtonLoadingStateEnum } from "./utils";
import { getButtonAppearance, getButtonIconPosition, getButtonShape, getButtonSize, getButtonStyle, getDisplayMode,} from "./utils";
import { 
  TransferProgressEvent, 
  IUploadResult} from "./PowerAutomateCloudFlows";

// File upload state with progress tracking
export interface IFileState {
  file: File;
  progress: number;
  status: 'pending' | 'uploading' | 'completed' | 'failed' | 'invalid';
  url?: string;
  error?: string;
  isValid?: boolean;
}

// State for files already in Azure storage
export interface IExistingFileState {
  name: string;
  size: number;
  url: string;
  lastModified: Date;
  isExisting: true; // Distinguishes from new files
}

// Admin configuration state for runtime property changes
export interface IAdminConfig {
  maxTotalFileSizeMB: number;
  allowedFileTypes: string;
}

export interface IFilesImportControlProps {
  buttonText: string;
  buttonIcon: string;
  buttonIconStyle: string;
  buttonAppearance: string;
  buttonAlign: string;
  buttonVisible: boolean;
  buttonFontWeight: string;
  buttonDisplayMode: string;
  buttonWidth: number;
  buttonHeight: number;
  buttonShowActionSpinner: boolean;
  buttonIconPosition: string;
  buttonShape: string;
  buttonButtonSize: string;
  buttonDisabledFocusable: boolean;
  buttonAllowMultipleFiles: boolean;
  buttonAllowedFileTypes: string;
  buttonAllowDropFiles: boolean;
  buttonAllowDropFilesText: string;
  maxTotalFileSizeMB: number;
  // Cloud Flow Configuration Properties
  cloudFlowUploadUrl?: string | null;
  cloudFlowListFilesUrl?: string | null;
  cloudFlowDeleteUrl?: string | null;
  cloudFlowDownloadUrl?: string | null;
  containerPath?: string | null;
  listFilesFolderName?: string | null;
  recordUid?: string | null;
  canvasAppCurrentTheme: Theme;
  context: ComponentFramework.Context<any>; // PCF Context for accessing localized resources
  onEvent: (event: any) => void;
}

export const FilesImportControl: React.FC<IFilesImportControlProps> = ({
  buttonText,
  buttonIcon,
  buttonIconStyle,
  buttonAppearance,
  buttonAlign,
  buttonVisible,
  buttonFontWeight,
  buttonDisplayMode,
  buttonWidth,
  buttonHeight,
  buttonShowActionSpinner,
  buttonIconPosition,
  buttonShape,
  buttonButtonSize,
  buttonDisabledFocusable,
  buttonAllowMultipleFiles,
  buttonAllowedFileTypes,
  buttonAllowDropFiles,
  buttonAllowDropFilesText,
  maxTotalFileSizeMB,
  cloudFlowUploadUrl,
  cloudFlowListFilesUrl,
  cloudFlowDeleteUrl,
  cloudFlowDownloadUrl,
  containerPath,
  listFilesFolderName,
  recordUid,
  canvasAppCurrentTheme,
  context,
  onEvent
}) => {
  
  // Localization helper function
  const getLocalizedString = (key: string, fallback: string): string => {
    try {
      const localizedValue = context.resources.getString(key);
      // Check if the value returned is the same as the key (meaning no localization found)
      return localizedValue !== key ? localizedValue : fallback;
    } catch (error) {
      console.warn(`Failed to get localized string for key '${key}':`, error);
      return fallback;
    }
  };
  
  // Use provided theme or default
  const _theme = canvasAppCurrentTheme?.fontFamilyBase?.trim()
    ? canvasAppCurrentTheme
    : webLightTheme;

  // Coca-Cola brand colors
  const cokeRed = '#F40009';
  
  // Create custom theme with Coca-Cola red for primary buttons
  const customTheme = {
    ..._theme,
    colorBrandBackground: cokeRed,
    colorBrandBackgroundHover: '#D40007',
    colorBrandBackgroundPressed: '#B40006',
    colorBrandForeground1: cokeRed,
    colorBrandForeground2: '#D40007'
  };

  // Button loading state management
  const [buttonLoadingState, setButtonLoadingState] = useState<ButtonLoadingStateEnum>(
    ButtonLoadingStateEnum.Initial
  );

  // Get button icon based on loading state
  const getButtonIcon = (): React.ReactNode => {
    if (buttonLoadingState === "loading") return <Spinner size="tiny" />;
    if (buttonLoadingState === "loaded") return <CheckmarkFilled />;
    return getIcon(buttonIcon, buttonIconStyle === "1");
  };

  // Common button properties
  const _buttonProps = {
    disabledFocusable: buttonDisabledFocusable,
    appearance: getButtonAppearance(buttonAppearance),
    iconPosition: getButtonIconPosition(buttonIconPosition),
    shape: getButtonShape(buttonShape),
    size: getButtonSize(buttonButtonSize),
    style: getButtonStyle(buttonWidth, buttonHeight, buttonAlign, buttonFontWeight),
    disabled: getDisplayMode(buttonLoadingState, buttonDisplayMode),
  };

  // File state management
  const [fileStates, setFileStates] = useState<IFileState[]>([]);
  const [existingFiles, setExistingFiles] = useState<IExistingFileState[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [uploadProgress, setUploadProgress] = useState<{ [fileName: string]: number }>({});
  const [showFileList, setShowFileList] = useState<boolean>(false);
  const [loadingExistingFiles, setLoadingExistingFiles] = useState<boolean>(false);
  const [isInitialLoad, setIsInitialLoad] = useState<boolean>(true);
  const [viewingFiles, setViewingFiles] = useState<{ [fileName: string]: boolean }>({});
  const [deletingFiles, setDeletingFiles] = useState<{ [fileName: string]: boolean }>({});
  const [showDeleteConfirmation, setShowDeleteConfirmation] = useState<{ [fileName: string]: boolean }>({});
  const [cumulativeFilesJSON, setCumulativeFilesJSON] = useState<any[]>([]); // Cumulative list of all uploaded files
  
  // Admin configuration state for runtime property changes
  const [adminConfig, setAdminConfig] = useState<IAdminConfig>(() => {
    // Try to load from localStorage first, fall back to props
    try {
      const saved = localStorage.getItem('fileUploadControlAdminConfig');
      if (saved) {
        const parsed = JSON.parse(saved);
        // Validate that saved config has the expected properties
        if (parsed.maxTotalFileSizeMB && parsed.allowedFileTypes !== undefined) {
          return parsed;
        }
      }
    } catch (error) {
      console.warn('Failed to load admin config from localStorage:', error);
    }
    
    // Fall back to props if no valid saved config
    return {
      maxTotalFileSizeMB: maxTotalFileSizeMB,
      allowedFileTypes: buttonAllowedFileTypes
    };
  });
  
  // Persist admin config changes to localStorage
  React.useEffect(() => {
    try {
      localStorage.setItem('fileUploadControlAdminConfig', JSON.stringify(adminConfig));
    } catch (error) {
      console.warn('Failed to save admin config to localStorage:', error);
    }
  }, [adminConfig]);
  
  // Reset admin config when props change (design-time override)
  React.useEffect(() => {
    setAdminConfig(prev => ({
      maxTotalFileSizeMB: maxTotalFileSizeMB,
      allowedFileTypes: buttonAllowedFileTypes
    }));
  }, [maxTotalFileSizeMB, buttonAllowedFileTypes]);
  
  // Get current configuration values (admin config takes precedence over props)
  const getCurrentConfig = () => ({
    maxTotalFileSizeMB: adminConfig.maxTotalFileSizeMB,
    allowedFileTypes: adminConfig.allowedFileTypes
  });
  
  // Generate localized secondary content text for allowed file types
  const getSecondaryContentText = (): string => {
    const config = getCurrentConfig();
    
    // If no file type restrictions or wildcard, don't show secondary content
    if (!config.allowedFileTypes || config.allowedFileTypes.trim() === '' || config.allowedFileTypes.trim() === '*') {
      return '';
    }
    
    // Parse allowed file types and clean them up for display
    const allowedTypes = config.allowedFileTypes
      .split(/[;,]/) // Support both semicolon and comma separators
      .map(type => type.trim().toLowerCase())
      .filter(type => type !== '' && type !== '*')
      .map(type => {
        // Remove dots for cleaner display (e.g., ".pdf" becomes "pdf")
        return type.startsWith('.') ? type.substring(1) : type;
      });
    
    if (allowedTypes.length === 0) {
      return '';
    }
    
    // Get localized "Allowed" prefix
    const allowedPrefix = getLocalizedString('AllowedPrefix', 'Allowed');
    
    // Join the types with commas
    const typesList = allowedTypes.join(',');
    
    return `${allowedPrefix} ${typesList}`;
  };
  
  const importFileRef = createRef<HTMLInputElement>();

  // Check Cloud Flow configuration
  const isCloudFlowConfigured = () => {
    return recordUid && recordUid.trim() !== '' && cloudFlowDownloadUrl && cloudFlowUploadUrl && cloudFlowListFilesUrl && cloudFlowDeleteUrl && containerPath;
  };

  // Get combined folder path using ListFilesFolderName and RecordUid
  const getCombinedFolderPath = () => {
    if (!listFilesFolderName || !recordUid) {
      return recordUid || listFilesFolderName || '';
    }
    return `${listFilesFolderName}/${recordUid}`;
  };

  // Check if download Cloud Flow is configured specifically
  const isDownloadCloudFlowConfigured = () => {
    return isCloudFlowConfigured() && cloudFlowDownloadUrl;
  };

  // Fetch existing files from Cloud Flow
  const loadExistingFiles = async () => {
    if (!isCloudFlowConfigured()) {
      // Clear existing files when cloud flow is not configured (e.g., when recordUid is empty)
      setExistingFiles([]);
      setShowFileList(false);
      setIsInitialLoad(false);
      onEvent({ 
        existingFiles: JSON.stringify([])
      });
      return;
    }

    setLoadingExistingFiles(true);
    try {
      // Create a simplified request payload
      const combinedFolderPath = getCombinedFolderPath();
      const requestPayload = {
        containerPath: containerPath!,
        folderName: combinedFolderPath
      };

      // Call the list files flow directly using fetch
      const response = await fetch(cloudFlowListFilesUrl!, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestPayload)
      });

      // Parse the response regardless of status code to check for files array
      let result;
      try {
        result = await response.json();
      } catch (parseError) {
        // If we can't parse JSON and response is not ok, throw the HTTP error
        if (!response.ok) {
          throw new Error(`List files failed: ${response.statusText}`);
        }
        throw new Error('Failed to parse response');
      }
      
      // Check if we have a valid files array (even if empty)
      const hasValidFilesArray = result.files && Array.isArray(result.files);
      
      // Only throw an error if:
      // 1. Response is not ok AND there's no valid files array, OR
      // 2. Success is false AND there's no valid files array
      if ((!response.ok || !result.success) && !hasValidFilesArray) {
        throw new Error(result.error || `List files failed: ${response.statusText}`);
      }

      // Handle files array - it may be empty if no files exist
      const files = result.files || [];
      const existingFileStates: IExistingFileState[] = files.map((file: any) => ({
        name: file.name,
        size: file.size,
        url: file.url,
        lastModified: new Date(file.lastModified),
        isExisting: true as const
      }));

      setExistingFiles(existingFileStates);
      
      if (existingFileStates.length > 0) {
        setShowFileList(true);
      } else if (isInitialLoad) {
        // Show file list with "no files" message for initial load even if no files exist
        setShowFileList(true);
      }

      // Notify parent about existing files
      onEvent({ 
        existingFiles: JSON.stringify(existingFileStates)
      });
    } catch (error) {
      console.error('Error loading existing files:', error);
      // Still show file list on error during initial load to show error state
      if (isInitialLoad) {
        setShowFileList(true);
      }
    } finally {
      setLoadingExistingFiles(false);
      setIsInitialLoad(false);
    }
  };

  // Load existing files on mount and prop changes
  React.useEffect(() => {
    // Reset initial load state when recordUid changes (new record selected)
    if (recordUid && recordUid.trim() !== '') {
      setIsInitialLoad(true);
    }
    
    loadExistingFiles();
    // Reset all file states when context changes (different record, container, etc.)
    setFileStates([]);
    setSelectedFiles([]);
    setUploadProgress({});
    setCumulativeFilesJSON([]);
    setShowDeleteConfirmation({});
    setDeletingFiles({});
    setViewingFiles({});
    onEvent({ 
      filesJSON: JSON.stringify([]),
      contextChanged: true
    });
  }, [cloudFlowListFilesUrl, containerPath, listFilesFolderName, recordUid]);

  // Validate file type against allowed formats and total size limit
  const validateFileType = (file: File, existingFiles: File[] = []): { isValid: boolean; error?: string } => {
    const config = getCurrentConfig();
    
    // First check file type restrictions
    if (!config.allowedFileTypes || config.allowedFileTypes.trim() === '' || config.allowedFileTypes.trim() === '*') {
      // No file type restrictions, continue to size check
    } else {
      // Parse allowed file types (semicolon/comma separated)
      const allowedTypes = config.allowedFileTypes
        .split(/[;,]/) // Support both semicolon and comma separators
        .map(type => type.trim().toLowerCase())
        .filter(type => type !== '' && type !== '*');

      if (allowedTypes.length > 0) {
        const fileName = file.name.toLowerCase();
        const fileType = file.type.toLowerCase();
        const fileExtension = fileName.substring(fileName.lastIndexOf('.'));

        // Check MIME types and extensions
        const isValidType = allowedTypes.some(allowedType => {
          // Check MIME type
          if (allowedType.includes('/')) {
            return fileType === allowedType || fileType.startsWith(allowedType.replace('*', ''));
          }
          // Check extension (with or without dot)
          const normalizedAllowedType = allowedType.startsWith('.') ? allowedType : `.${allowedType}`;
          return fileExtension === normalizedAllowedType;
        });

        if (!isValidType) {
          return {
            isValid: false,
            error: `File type not allowed. Allowed: ${config.allowedFileTypes}`
          };
        }
      }
    }

    // Check total file size limit
    if (config.maxTotalFileSizeMB && config.maxTotalFileSizeMB > 0) {
      const maxSizeBytes = config.maxTotalFileSizeMB * 1024 * 1024; // Convert MB to bytes
      
      // Calculate total size of existing pending/valid files plus the new file
      const existingPendingFiles = fileStates.filter(fs => fs.isValid && fs.status !== 'failed').map(fs => fs.file);
      const allExistingFiles = [...existingPendingFiles, ...existingFiles];
      const totalExistingSize = allExistingFiles.reduce((sum, f) => sum + f.size, 0);
      const totalNewSize = totalExistingSize + file.size;

      if (totalNewSize > maxSizeBytes) {
        const totalSizeMB = (totalNewSize / (1024 * 1024)).toFixed(2);
        const currentSizeMB = (totalExistingSize / (1024 * 1024)).toFixed(2);
        const fileSizeMB = (file.size / (1024 * 1024)).toFixed(2);
        return {
          isValid: false,
          error: `Adding this file (${fileSizeMB} MB) would exceed the ${config.maxTotalFileSizeMB} MB limit. Current total: ${currentSizeMB} MB`
        };
      }
    }

    return { isValid: true };
  };

  // Add files to state with validation
  const addFilesToState = (files: File[]) => {
    const newFileStates: IFileState[] = [];
    let currentBatchSize = 0;
    
    // Calculate current total size of valid pending files
    const existingPendingFiles = fileStates.filter(fs => fs.isValid && fs.status !== 'failed').map(fs => fs.file);
    const existingPendingSize = existingPendingFiles.reduce((sum, f) => sum + f.size, 0);
    
    for (const file of files) {
      // For batch validation, include previously processed files in this batch
      const filesInCurrentBatch = newFileStates.filter(fs => fs.isValid).map(fs => fs.file);
      const validation = validateFileType(file, filesInCurrentBatch);
      
      newFileStates.push({
        file,
        progress: 0,
        status: validation.isValid ? 'pending' as const : 'invalid' as const,
        isValid: validation.isValid,
        error: validation.error
      });
      
      if (validation.isValid) {
        currentBatchSize += file.size;
      }
    }
    
    setFileStates(prev => [...prev, ...newFileStates]);
    setSelectedFiles(prev => [...prev, ...files]);
    setShowFileList(true);
  };

  // Update specific file state
  const updateFileState = (fileName: string, updates: Partial<IFileState>) => {
    setFileStates(prev => prev.map(fileState => 
      fileState.file.name === fileName 
        ? { ...fileState, ...updates }
        : fileState
    ));
  };

  // Remove file from state (only pending/invalid files)
  const removeFile = (fileName: string) => {
    // Only allow removal of pending or invalid files
    const fileState = fileStates.find(fs => fs.file.name === fileName);
    if (fileState && (fileState.status === 'uploading' || fileState.status === 'completed' || fileState.status === 'failed')) {
      console.warn(`Cannot remove file ${fileName} - file is ${fileState.status}`);
      return;
    }

    setFileStates(prev => {
      const newFileStates = prev.filter(fileState => fileState.file.name !== fileName);
      
      // Hide file list if no files remain
      if (newFileStates.length === 0 && existingFiles.length === 0) {
        setShowFileList(false);
      }
      
      return newFileStates;
    });
    setSelectedFiles(prev => prev.filter(file => file.name !== fileName));
    
    // Only manage cumulative JSON when RecordUid is empty
    if (!recordUid || recordUid.trim() === '') {
      // Remove from cumulative JSON if it was already processed
      setCumulativeFilesJSON(prev => {
        const updatedCumulative = prev.filter(file => file.name !== fileName);
        // Notify parent about the updated cumulative files
        onEvent({ 
          filesJSON: JSON.stringify(updatedCumulative),
          fileRemovedFromCumulative: fileName
        });
        return updatedCumulative;
      });
    }
  };

  // Check if any upload is in progress
  const isUploadInProgress = () => {
    return fileStates.some(fileState => fileState.status === 'uploading');
  };

  // Clear all new files (preserve existing files)
  const clearNewFiles = () => {
    setFileStates([]);
    setSelectedFiles([]);
    setUploadProgress({});
    setShowDeleteConfirmation({});
    setDeletingFiles({});
    
    // Only manage cumulative JSON when RecordUid is empty
    if (!recordUid || recordUid.trim() === '') {
      setCumulativeFilesJSON([]); // Clear cumulative files as well
      
      // Notify parent that cumulative files have been cleared
      onEvent({ 
        filesJSON: JSON.stringify([]),
        cumulativeFilesCleared: true
      });
    }
    
    // Hide file list if no existing files remain
    if (existingFiles.length === 0) {
      setShowFileList(false);
    }
  };

  // Delete existing file from Cloud Flow
  const handleDeleteExistingFile = async (fileName: string) => {
    setDeletingFiles(prev => ({ ...prev, [fileName]: true }));

    try {
      if (!isCloudFlowConfigured()) {
        throw new Error('Cloud Flow configuration is missing');
      }

      // Create request payload based on COCA postman collection
      const combinedFolderPath = getCombinedFolderPath();
      const requestPayload = {
        fileName: fileName,
        containerPath: containerPath!,
        folderName: combinedFolderPath || undefined
      };

      // Call the delete flow directly using fetch
      const response = await fetch(cloudFlowDeleteUrl!, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestPayload)
      });

      if (!response.ok) {
        throw new Error(`Delete failed: ${response.statusText}`);
      }

      const result = await response.json();
      
      if (result.success) {
        // Remove from UI state
        setExistingFiles(prev => prev.filter(file => file.name !== fileName));
        
        // Refresh existing files list
        await loadExistingFiles();
        
        // Notify parent about deletion
        onEvent({
          fileDeleted: fileName,
          existingFiles: JSON.stringify(existingFiles.filter(file => file.name !== fileName))
        });
        
        console.log(`File ${fileName} deleted successfully`);
      } else {
        throw new Error(result.error || 'Failed to delete file');
      }
    } catch (error) {
      console.error('Error deleting file:', error);
      alert(`Failed to delete file: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setDeletingFiles(prev => ({ ...prev, [fileName]: false }));
      setShowDeleteConfirmation(prev => ({ ...prev, [fileName]: false }));
    }
  };

  // Show delete confirmation buttons
  const showDeleteConfirmationButtons = (fileName: string) => {
    setShowDeleteConfirmation(prev => ({ ...prev, [fileName]: true }));
  };

  // Cancel delete confirmation
  const cancelDeleteConfirmation = (fileName: string) => {
    setShowDeleteConfirmation(prev => ({ ...prev, [fileName]: false }));
  };

  // Confirm and execute delete
  const confirmDelete = async (fileName: string) => {
    await handleDeleteExistingFile(fileName);
  };

  // Download existing file using the new download Cloud Flow
  const handleViewExistingFile = async (fileName: string) => {
    // Fallback to direct URL if Cloud Flow not configured for download
    if (!cloudFlowDownloadUrl) {
      const existingFile = existingFiles.find(f => f.name === fileName);
      if (existingFile?.url) {
        window.open(existingFile.url, '_blank', 'noopener,noreferrer');
        onEvent({
          fileViewed: fileName,
          viewUrl: existingFile.url,
          method: 'direct'
        });
      } else {
        alert('No URL available for this file');
      }
      return;
    }

    if (!isDownloadCloudFlowConfigured()) {
      alert('Cloud Flow configuration for download is missing');
      return;
    }

    // Set loading state for this specific file
    setViewingFiles(prev => ({ ...prev, [fileName]: true }));

    try {
      // Create request payload based on the new download flow schema
      const combinedFolderPath = getCombinedFolderPath();
      const requestPayload = {
        storageAccountName: containerPath!,
        filePath: combinedFolderPath ? `${combinedFolderPath}/${fileName}` : fileName,
        fileName: fileName
      };

      // Call the download flow directly using fetch
      const response = await fetch(cloudFlowDownloadUrl!, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestPayload)
      });

      if (!response.ok) {
        throw new Error(`Download file failed: ${response.statusText}`);
      }

      const result = await response.json();
      
      // Handle the response structure from the new download flow
      let fileContent: string | null = null;
      let contentType: string | null = null;
      let downloadFileName: string | null = null;
      
      // Check different possible response structures
      if (result.fileContent && result.contentType && result.fileName) {
        // Direct response structure
        fileContent = result.fileContent;
        contentType = result.contentType;
        downloadFileName = result.fileName;
      } else if (result.body && result.body.fileContent && result.body.contentType && result.body.fileName) {
        // Power Automate response with body wrapper
        fileContent = result.body.fileContent;
        contentType = result.body.contentType;
        downloadFileName = result.body.fileName;
      }
      
      if (fileContent && contentType && downloadFileName) {
        // Create downloadable link using the base64 file content
        const link = document.createElement('a');
        link.href = `data:${contentType};base64,${fileContent}`;
        link.download = downloadFileName;
        
        // Trigger the download
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        // Notify parent about download action
        onEvent({
          fileDownloaded: downloadFileName,
          fileSize: result.fileSize || result.body?.fileSize || 0,
          contentType: contentType,
          method: 'cloudflow'
        });
        
        console.log(`Downloaded file ${downloadFileName} successfully`);
      } else {
        // Log the actual response structure for debugging
        console.error('Unexpected response structure:', result);
        throw new Error(result.error || result.body?.error || 'Failed to download file - missing file content in response');
      }
    } catch (error) {
      console.error('Error downloading file:', error);
      alert(`Failed to download file: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      // Clear loading state for this specific file
      setViewingFiles(prev => ({ ...prev, [fileName]: false }));
    }
  };

  // Format bytes to readable size
  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
  };

  // Calculate current total file size of pending/valid new files
  const getCurrentTotalFileSize = (): number => {
    const validPendingFiles = fileStates.filter(fs => fs.isValid && fs.status !== 'failed');
    return validPendingFiles.reduce((total, fs) => total + fs.file.size, 0);
  };

  // Read file as data URL
  const readFile = (file: File): Promise<string | null> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      // Handle read errors
      reader.onerror = () => {
        console.error(`Error reading file: ${file.name}`);
        reject(null);
      };

      // Return result when reading completes
      reader.onload = () => {
        resolve(reader.result as string);
      };

      reader.readAsDataURL(file);
    });
  };

  // Progress simulation for fast uploads
  const progressSimulationIntervals = React.useRef<{ [fileName: string]: NodeJS.Timeout }>({});

  const startProgressSimulation = (fileName: string, fileSize: number) => {
    let simulatedProgress = 0;
    progressSimulationIntervals.current[fileName] = setInterval(() => {
      simulatedProgress = Math.min(simulatedProgress + Math.random() * 15, 95); // Cap at 95%
      updateFileState(fileName, { progress: Math.floor(simulatedProgress) });
    }, 200); // Update every 200ms
  };

  const stopProgressSimulation = (fileName: string) => {
    if (progressSimulationIntervals.current[fileName]) {
      clearInterval(progressSimulationIntervals.current[fileName]);
      delete progressSimulationIntervals.current[fileName];
    }
  };

  // Clean up all progress simulations
  React.useEffect(() => {
    return () => {
      Object.values(progressSimulationIntervals.current).forEach(interval => {
        clearInterval(interval);
      });
    };
  }, []);

  // Upload files to Cloud Flow with progress tracking
  const uploadToCloudFlow = async (files: File[]) => {
    if (!isCloudFlowConfigured()) {
      throw new Error('Cloud Flow configuration is missing');
    }

    // Initialize upload state
    files.forEach(file => {
      updateFileState(file.name, { status: 'uploading', progress: 0 });
    });

    if (buttonAllowMultipleFiles) {
      // Multiple file upload - process each file
      const results: IUploadResult[] = [];
      
      for (const file of files) {
        startProgressSimulation(file.name, file.size);
        
        try {
          // Convert file to base64
          const reader = new FileReader();
          const base64Content = await new Promise<string>((resolve, reject) => {
            reader.onload = () => {
              if (reader.result) {
                const base64 = (reader.result as string).split(',')[1];
                resolve(base64);
              } else {
                reject(new Error('Failed to read file'));
              }
            };
            reader.onerror = () => reject(reader.error);
            reader.readAsDataURL(file);
          });

          // Create request payload based on COCA postman collection structure
          const combinedFolderPath = getCombinedFolderPath();
          const requestPayload = {
            fileName: file.name,
            fileContent: base64Content,
            containerPath: containerPath!,
            folderName: combinedFolderPath,
            fileSize: file.size,
            contentType: file.type || 'application/octet-stream'
          };

          // Call the upload flow directly using fetch
          const response = await fetch(cloudFlowUploadUrl!, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify(requestPayload)
          });

          if (!response.ok) {
            throw new Error(`Upload failed: ${response.statusText}`);
          }

          const result = await response.json();
          
          stopProgressSimulation(file.name);
          
          if (result.success) {
            updateFileState(file.name, { progress: 100 });
            results.push({
              fileName: file.name,
              url: result.url || '',
              success: true,
              flowRunId: result.flowRunId
            });
          } else {
            results.push({
              fileName: file.name,
              url: '',
              success: false,
              error: result.error || 'Upload failed'
            });
          }
        } catch (error) {
          stopProgressSimulation(file.name);
          results.push({
            fileName: file.name,
            url: '',
            success: false,
            error: error instanceof Error ? error.message : String(error)
          });
        }
      }
      
      // Update file states with results
      results.forEach(result => {
        updateFileState(result.fileName, {
          status: result.success ? 'completed' : 'failed',
          url: result.success ? result.url : undefined,
          error: result.success ? undefined : result.error,
          progress: result.success ? 100 : 0
        });
      });
      
      return results;
    } else {
      // Single file upload
      startProgressSimulation(files[0].name, files[0].size);
      
      try {
        // Convert file to base64
        const reader = new FileReader();
        const base64Content = await new Promise<string>((resolve, reject) => {
          reader.onload = () => {
            if (reader.result) {
              const base64 = (reader.result as string).split(',')[1];
              resolve(base64);
            } else {
              reject(new Error('Failed to read file'));
            }
          };
          reader.onerror = () => reject(reader.error);
          reader.readAsDataURL(files[0]);
        });

        // Create request payload
        const combinedFolderPath = getCombinedFolderPath();
        const requestPayload = {
          fileName: files[0].name,
          fileContent: base64Content,
          containerPath: containerPath!,
          folderName: combinedFolderPath,
          fileSize: files[0].size,
          contentType: files[0].type || 'application/octet-stream'
        };

        // Call the upload flow directly using fetch
        const response = await fetch(cloudFlowUploadUrl!, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(requestPayload)
        });

        if (!response.ok) {
          throw new Error(`Upload failed: ${response.statusText}`);
        }

        const result = await response.json();
        
        stopProgressSimulation(files[0].name);
        
        if (result.success) {
          updateFileState(files[0].name, { progress: 100 });
          const uploadResult: IUploadResult = {
            fileName: files[0].name,
            url: result.url || '',
            success: true,
            flowRunId: result.flowRunId
          };
          
          updateFileState(files[0].name, {
            status: 'completed',
            url: uploadResult.url,
            progress: 100
          });
          
          return [uploadResult];
        } else {
          const uploadResult: IUploadResult = {
            fileName: files[0].name,
            url: '',
            success: false,
            error: result.error || 'Upload failed'
          };
          
          updateFileState(files[0].name, {
            status: 'failed',
            error: uploadResult.error,
            progress: 0
          });
          
          return [uploadResult];
        }
      } catch (error) {
        stopProgressSimulation(files[0].name);
        const uploadResult: IUploadResult = {
          fileName: files[0].name,
          url: '',
          success: false,
          error: error instanceof Error ? error.message : String(error)
        };
        
        updateFileState(files[0].name, {
          status: 'failed',
          error: uploadResult.error,
          progress: 0
        });
        
        return [uploadResult];
      }
    }
  };

  // Process files as JSON
  const processFilesAsJSON = async (files: File[]) => {
    const filesArray = await Promise.all(
      files.map(async (file) => {
        const fileContent = await readFile(file);
        return {
          name: file.name,
          size: file.size,
          contentBytes: fileContent || "",
        };
      })
    );

    const jsonString = JSON.stringify(filesArray);
    return jsonString;
  };

  // Process files by uploading to Cloud Flow or creating JSON
  const processFiles = async (files: File[]) => {
    if (!files || files.length === 0) return;

    try {
      if (buttonShowActionSpinner) {
        setButtonLoadingState(ButtonLoadingStateEnum.Loading);
      }
      
      if (isCloudFlowConfigured()) {
        // Upload to Cloud Flow
        onEvent({ uploadStatus: "InProgress" });
        
        const uploadResults = await uploadToCloudFlow(files);
        const hasFailures = uploadResults.some((result: IUploadResult) => !result.success);
        
        // Reload existing files and clean up after successful uploads
        if (uploadResults.some((result: IUploadResult) => result.success)) {
          await loadExistingFiles();
          
          // Remove successfully uploaded files from state to avoid duplicates
          const successfulUploads = uploadResults.filter((result: IUploadResult) => result.success).map((result: IUploadResult) => result.fileName);
          setFileStates(prev => prev.filter(fileState => !successfulUploads.includes(fileState.file.name)));
          setSelectedFiles(prev => prev.filter(file => !successfulUploads.includes(file.name)));
          
          // Clear progress for successful uploads
          setUploadProgress(prev => {
            const newProgress = { ...prev };
            successfulUploads.forEach((fileName: string) => {
              delete newProgress[fileName];
            });
            return newProgress;
          });
        }
        
        // Notify parent with upload results (no cumulative JSON for cloud flow uploads)
        onEvent({ 
          uploadResults: JSON.stringify(uploadResults),
          uploadStatus: hasFailures ? "Failed" : "Completed",
          filesJSON: JSON.stringify(uploadResults.map((result: IUploadResult) => ({
            name: result.fileName,
            url: result.url,
            success: result.success,
            error: result.error
          })))
        });
      } else {
        // Fallback to JSON processing - only use cumulative JSON when RecordUid is empty
        const jsonString = await processFilesAsJSON(files);
        const currentBatchFiles = JSON.parse(jsonString);
        
        // Mark all files as completed for JSON mode
        files.forEach(file => {
          updateFileState(file.name, { status: 'completed', progress: 100 });
        });
        
        // Only maintain cumulative JSON when RecordUid is empty
        if (!recordUid || recordUid.trim() === '') {
          // Update cumulative files JSON
          setCumulativeFilesJSON(prev => [...prev, ...currentBatchFiles]);
          const updatedCumulativeFiles = [...cumulativeFilesJSON, ...currentBatchFiles];
          
          onEvent({ 
            filesJSON: JSON.stringify(updatedCumulativeFiles), // Send cumulative files
            uploadStatus: "Completed"
          });
        } else {
          // When RecordUid is present, just send current batch
          onEvent({ 
            filesJSON: jsonString,
            uploadStatus: "Completed"
          });
          
          // Clear the file states after processing since we're not maintaining cumulative state
          setFileStates(prev => prev.filter(fileState => !files.some(file => file.name === fileState.file.name)));
          setSelectedFiles(prev => prev.filter(file => !files.some(f => f.name === file.name)));
        }
      }

      // Set button to loaded state
      if (buttonShowActionSpinner) {
        setButtonLoadingState(ButtonLoadingStateEnum.Loaded);
      }
    } catch (error) {
      console.error("Error processing files:", error);
      setButtonLoadingState(ButtonLoadingStateEnum.Initial);
      
      // Stop all progress simulations
      files.forEach(file => stopProgressSimulation(file.name));
      
      // Mark all files as failed
      files.forEach(file => {
        updateFileState(file.name, { 
          status: 'failed', 
          error: error instanceof Error ? error.message : String(error),
          progress: 0 
        });
      });
      
      onEvent({ 
        uploadStatus: "Failed",
        uploadResults: JSON.stringify([{
          fileName: "Unknown",
          url: "",
          success: false,
          error: error instanceof Error ? error.message : String(error)
        }])
      });
    }
  };

  // Handle file input change event
  const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      // Add files to state without auto-upload
      const filesArray = Array.from(event.target.files);
      addFilesToState(filesArray);
      event.target.value = "";
    }
  };

  // Handle upload button click
  const handleButtonClick = () => {
    setButtonLoadingState(ButtonLoadingStateEnum.Initial);
    importFileRef.current?.click();
  };

  // Upload all pending valid files
  const handleUploadAll = async () => {
    const pendingValidFiles = fileStates
      .filter(fs => fs.status === 'pending' && fs.isValid)
      .map(fs => fs.file);
    
    if (pendingValidFiles.length > 0) {
      await processFiles(pendingValidFiles);
    }
  };

  // Drag & drop state management
  const [isDragging, setIsDragging] = React.useState<boolean>(false);
  const dropZoneRef = React.useRef<HTMLDivElement>(null);

  // Handle drag enter event
  const onDragEnter = (event: React.DragEvent) => {
    if (buttonDisplayMode !== '0') return; // Only in edit mode
    event.preventDefault();
    setIsDragging(true);
  };

  // Handle drag over event
  const onDragOver = (event: React.DragEvent) => {
    if (buttonDisplayMode !== '0') return; // Only in edit mode
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
  };

  // Handle drag leave event
  const onDragLeave = (event: React.DragEvent) => {
    if (buttonDisplayMode !== '0') return; // Only in edit mode
    if (!dropZoneRef.current?.contains(event.relatedTarget as Node)) {
      setIsDragging(false);
    }
  };

  // Handle file drop event
  const onDrop = async (event: React.DragEvent) => {
    if (buttonDisplayMode !== '0') return; // Only in edit mode
    event.preventDefault();
    setIsDragging(false);

    if (event.dataTransfer.files) {
      // Add files to state without auto-upload
      const filesArray = Array.from(event.dataTransfer.files);
      addFilesToState(filesArray);
    }
  };

  return (
    <FluentProvider theme={customTheme}>
      <div style={{ 
        width: '100%', 
        minWidth: '400px',
        height: 'auto',
        minHeight: '400px',
        display: 'flex',
        flexDirection: 'column'
      }}>
        {/* Admin Configuration Mode */}
        {buttonDisplayMode === '3' ? (
          <div style={{
            padding: '20px',
            border: '1px solid #e1dfdd',
            borderRadius: '8px',
            backgroundColor: 'white',
            margin: '10px'
          }}>
            <div style={{ marginBottom: '20px', borderBottom: '1px solid #e1dfdd', paddingBottom: '10px' }}>
              <Caption1 style={{ color: '#605e5c', marginTop: '4px' }}>
                {getLocalizedString('ConfigureRuntimeProperties', 'Configure runtime properties for the file upload control')}
              </Caption1>
            </div>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
              {/* Max File Size Configuration */}
              <div>
                <Text style={{ fontWeight: 'semibold', marginBottom: '8px', display: 'block' }}>
                  {getLocalizedString('MaxTotalFileSize', 'Maximum Total File Size (MB)')}
                </Text>
                <input
                  type="number"
                  min="1"
                  max="100"
                  value={adminConfig.maxTotalFileSizeMB}
                  onChange={(e) => {
                    const value = e.target.value;
                    setAdminConfig(prev => ({
                      ...prev,
                      maxTotalFileSizeMB: parseInt(value) || 1
                    }));
                  }}
                  style={{
                    width: '100px',
                    padding: '8px 12px',
                    border: '1px solid #d1d1d1',
                    borderRadius: '4px',
                    fontSize: '14px'
                  }}
                />
                <Caption1 style={{ color: '#605e5c', marginTop: '4px', display: 'block' }}>
                  {getLocalizedString('Current', 'Current')}: {adminConfig.maxTotalFileSizeMB} MB ({getLocalizedString('Original', 'Original')}: {maxTotalFileSizeMB} MB)
                </Caption1>
              </div>

              {/* Allowed File Types Configuration */}
              <div>
                <Text style={{ fontWeight: 'semibold', marginBottom: '8px', display: 'block' }}>
                  {getLocalizedString('AllowedFileTypes', 'Allowed File Types')}
                </Text>
                <input
                  type="text"
                  value={adminConfig.allowedFileTypes}
                  onChange={(e) => {
                    const value = e.target.value;
                    setAdminConfig(prev => ({
                      ...prev,
                      allowedFileTypes: value
                    }));
                  }}
                  placeholder={getLocalizedString('FileTypePlaceholder', 'e.g., .pdf;.docx;.txt or * for all types')}
                  style={{
                    width: '100%',
                    maxWidth: '400px',
                    padding: '8px 12px',
                    border: '1px solid #d1d1d1',
                    borderRadius: '4px',
                    fontSize: '14px'
                  }}
                />
                <Caption1 style={{ color: '#605e5c', marginTop: '4px', display: 'block' }}>
                  {getLocalizedString('FileTypeInstructions', 'Use semicolons (;) or commas (,) to separate multiple types. Use * to allow all types.')}
                  <br />
                  {getLocalizedString('OriginalSetting', 'Original setting')}: "{buttonAllowedFileTypes}"
                </Caption1>
              </div>

              {/* Action Buttons */}
              <div style={{ marginTop: '16px', display: 'flex', gap: '12px' }}>
                <Button
                  appearance="primary"
                  onClick={() => {
                    // Reset to original values
                    setAdminConfig({
                      maxTotalFileSizeMB: maxTotalFileSizeMB,
                      allowedFileTypes: buttonAllowedFileTypes
                    });
                  }}
                >
                  {getLocalizedString('ResetToOriginal', 'Reset to Original')}
                </Button>
                <Button
                  appearance="secondary"
                  onClick={() => {
                    // Since we can't change the display mode directly, we'll just show a message
                    //alert('Admin configuration updated. Change the Display Mode property to exit admin mode.');
                    //don't do anything because we need to use a toggle button to switch between edit and admin mode
                  }}
                >
                  {getLocalizedString('Apply', 'Apply')}
                </Button>
              </div>

            </div>
          </div>
        ) : (
          <>
        {/* Upload Area */}
        <div
          ref={dropZoneRef}
          onDragEnter={buttonDisplayMode === '0' ? onDragEnter : undefined}
          onDragOver={buttonDisplayMode === '0' ? onDragOver : undefined}
          onDragLeave={buttonDisplayMode === '0' ? onDragLeave : undefined}
          onDrop={buttonDisplayMode === '0' ? onDrop : undefined}
          style={{
            border: isDragging && buttonAllowDropFiles && buttonDisplayMode === '0'
              ? `2px dashed ${cokeRed}` 
              : '1px solid #e1dfdd',
            borderRadius: '8px',
            padding: '20px',
            textAlign: 'center',
            backgroundColor: isDragging && buttonAllowDropFiles && buttonDisplayMode === '0'
              ? `${cokeRed}15` 
              : '#fafafa',
            /*cursor: 'pointer',*/
            transition: 'all 0.3s ease',
            minHeight: '120px',
            display: 'flex',
            flexDirection: 'column',
            justifyContent: 'center',
            alignItems: 'center',
            position: 'relative',
            flexShrink: 0
          }}
          /*onClick={handleButtonClick}*/
        >
          {buttonVisible && buttonDisplayMode === '0' ? (
            (() => {
              const secondaryContent = getSecondaryContentText();
              return secondaryContent ? (
                <CompoundButton 
                  // @ts-ignore - Icon type issue workaround
                  icon={getButtonIcon()} 
                  onClick={(e) => {
                    e.stopPropagation();
                    handleButtonClick();
                  }}
                  {..._buttonProps} 
                  secondaryContent={secondaryContent}
                  style={{
                    ..._buttonProps.style,
                    marginBottom: '10px'
                  }}
                >
                  {buttonText}
                </CompoundButton>
              ) : (
                <Button 
                  // @ts-ignore - Icon type issue workaround
                  icon={getButtonIcon()} 
                  onClick={(e) => {
                    e.stopPropagation();
                    handleButtonClick();
                  }}
                  {..._buttonProps}
                  style={{
                    ..._buttonProps.style,
                    marginBottom: '10px'
                  }}
                >
                  {buttonText}
                </Button>
              );
            })()
          ) : buttonDisplayMode === '2' ? (
            <Body1 style={{ color: '#605e5c', marginBottom: '10px' }}>
              {existingFiles.length > 0 ? `${getLocalizedString('ViewingFilesIn', 'Viewing files in')} ${getCombinedFolderPath() || getLocalizedString('Container', 'container')}` : getLocalizedString('NoFilesFound', 'No files found')}
            </Body1>
          ) : null}
          
          {buttonAllowDropFiles && buttonDisplayMode === '0' && (
            <Body1 style={{ color: '#605e5c', marginTop: '8px' }}>
              {isDragging ? buttonAllowDropFilesText : getLocalizedString('OrDropFilesHere', 'Or drop files here')}
            </Body1>
          )}
          
          {/* Version Label */}
          <Caption2 style={{ 
            position: 'absolute', 
            bottom: '8px', 
            right: '8px', 
            color: '#8a8886', 
            fontSize: '10px',
            opacity: 0.7
          }}>
            v2.0.23
          </Caption2>
        </div>

        {/* File Size Limitations - Below the upload area */}
        {buttonDisplayMode === '0' && getCurrentConfig().maxTotalFileSizeMB && getCurrentConfig().maxTotalFileSizeMB > 0 && (
          <Caption1 style={{ 
            color: '#605e5c', 
            fontSize: '12px', 
            marginTop: '8px', 
            display: 'block',
            textAlign: 'left',
            paddingLeft: '4px'
          }}>
            {getLocalizedString('MaximumTotalSize', 'Maximum total size')}: {getCurrentConfig().maxTotalFileSizeMB} MB
            {fileStates.length > 0 && ` • ${getLocalizedString('CurrentFiles', 'Current files')}: ${formatFileSize(getCurrentTotalFileSize())}`}
          </Caption1>
        )}

        {/* File List Display */}
        {(showFileList || (loadingExistingFiles && isInitialLoad)) && (
          <div style={{
            marginTop: '16px',
            border: '1px solid #e1dfdd',
            borderRadius: '8px',
            backgroundColor: 'white',
            display: 'flex',
            flexDirection: 'column',
            flex: '1 1 auto',
            minHeight: '200px',
            maxHeight: '600px'
          }}>
            {/* Header with file count and actions */}
            <div style={{
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              padding: '12px 16px',
              borderBottom: '1px solid #e1dfdd',
              flexShrink: 0
            }}>
              <Text weight="semibold">
                {loadingExistingFiles && isInitialLoad ? (
                  <Skeleton>
                    <SkeletonItem style={{ width: '200px', height: '20px' }} />
                  </Skeleton>
                ) : (
                  <>
                    {existingFiles.length > 0 && fileStates.length === 0
                      ? `${existingFiles.length} ${getLocalizedString('ExistingFiles', 'existing file(s)')}`
                      : fileStates.filter(fs => fs.status === 'completed').length > 0 
                      ? `${fileStates.filter(fs => fs.status === 'completed').length} ${getLocalizedString('FilesUploaded', 'file(s) uploaded')}${existingFiles.length > 0 ? `, ${existingFiles.length} ${getLocalizedString('Existing', 'existing')}` : ''}${(!recordUid || recordUid.trim() === '') && cumulativeFilesJSON.length > 0 ? ` • ${cumulativeFilesJSON.length} ${getLocalizedString('TotalInBatch', 'total in batch')}` : ''}`
                      : fileStates.some(fs => fs.status === 'invalid')
                      ? `${fileStates.filter(fs => fs.isValid).length} ${getLocalizedString('Valid', 'valid')}, ${fileStates.filter(fs => !fs.isValid).length} ${getLocalizedString('Invalid', 'invalid')} ${getLocalizedString('ExistingFiles', 'file(s)')}${existingFiles.length > 0 ? `, ${existingFiles.length} ${getLocalizedString('Existing', 'existing')}` : ''}${(!recordUid || recordUid.trim() === '') && cumulativeFilesJSON.length > 0 ? ` • ${cumulativeFilesJSON.length} ${getLocalizedString('TotalInBatch', 'total in batch')}` : ''}`
                      : `${fileStates.length} ${getLocalizedString('FilesReady', 'file(s) ready')}${existingFiles.length > 0 ? `, ${existingFiles.length} ${getLocalizedString('Existing', 'existing')}` : ''}${(!recordUid || recordUid.trim() === '') && cumulativeFilesJSON.length > 0 ? ` • ${cumulativeFilesJSON.length} ${getLocalizedString('TotalInBatch', 'total in batch')}` : ''}`}
                    {loadingExistingFiles && ` • ${getLocalizedString('Loading', 'Loading...')}`}
                  </>
                )}
              </Text>
              <div style={{ display: 'flex', gap: '8px' }}>
                {!loadingExistingFiles && fileStates.some(fs => fs.status === 'pending' && fs.isValid) && buttonDisplayMode === '0' && (
                  <Button
                    appearance="primary"
                    size="small"
                    onClick={handleUploadAll}
                    disabled={buttonLoadingState === 'loading' || !fileStates.some(fs => fs.status === 'pending' && fs.isValid)}
                  >
                    {getLocalizedString('UploadAll', 'Upload All')}
                  </Button>
                )}
                {!loadingExistingFiles && fileStates.length > 0 && !((!recordUid || recordUid.trim() === '') && cumulativeFilesJSON.length > 0 && fileStates.every(fs => fs.status === 'completed')) && (
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<DeleteRegular />}
                    onClick={clearNewFiles}
                    disabled={isUploadInProgress()}
                    title={fileStates.some(fs => fs.status === 'completed') ? getLocalizedString('ClearUploadedFilesTooltip', "Clear uploaded files (they're now in existing files)") : getLocalizedString('ClearNewFilesTooltip', "Clear new files")}
                  >
                    {getLocalizedString('ClearNewFiles', 'Clear New Files')}
                  </Button>
                )}
              </div>
            </div>

            {/* File items container */}
            <div style={{ 
              flex: '1 1 auto',
              overflowY: 'auto',
              minHeight: '100px',
              maxHeight: '500px'
            }}>
              {loadingExistingFiles && isInitialLoad ? (
                /* Skeleton Loader for Initial Load */
                <div style={{ padding: '16px' }}>
                  <div style={{
                    padding: '8px 0',
                    borderBottom: '1px solid #e1dfdd',
                    marginBottom: '12px'
                  }}>
                    <Skeleton>
                      <SkeletonItem style={{ width: '180px', height: '16px', marginBottom: '4px' }} />
                    </Skeleton>
                  </div>
                  {/* Skeleton for 3 file items */}
                  {[1, 2, 3].map((index) => (
                    <div key={`skeleton-${index}`} style={{
                      display: 'flex',
                      alignItems: 'center',
                      padding: '12px 0',
                      borderBottom: '1px solid #f3f2f1'
                    }}>
                      <div style={{ flex: 1 }}>
                        <Skeleton>
                          <SkeletonItem style={{ width: '200px', height: '18px', marginBottom: '6px' }} />
                          <SkeletonItem style={{ width: '140px', height: '14px' }} />
                        </Skeleton>
                      </div>
                      <div style={{ display: 'flex', gap: '8px' }}>
                        <Skeleton>
                          <SkeletonItem style={{ width: '60px', height: '32px' }} />
                        </Skeleton>
                        <Skeleton>
                          <SkeletonItem style={{ width: '60px', height: '32px' }} />
                        </Skeleton>
                      </div>
                    </div>
                  ))}
                </div>
              ) : existingFiles.length === 0 && fileStates.length === 0 && !loadingExistingFiles ? (
                /* No Files Message */
                <div style={{
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  justifyContent: 'center',
                  padding: '40px 20px',
                  textAlign: 'center',
                  minHeight: '150px'
                }}>
                  <Text style={{ 
                    color: '#605e5c', 
                    fontSize: '16px',
                    marginBottom: '8px'
                  }}>
                    {getLocalizedString('NoFilesExistInDirectory', 'No files exist in this directory')}
                  </Text>
                  <Caption1 style={{ color: '#8a8886' }}>
                    {getCombinedFolderPath() ? `${getLocalizedString('Directory', 'Directory')}: ${getCombinedFolderPath()}` : getLocalizedString('UploadFilesToGetStarted', 'Upload files to get started')}
                  </Caption1>
                </div>
              ) : (
                <>
                  {/* Existing Files Section */}
                  {existingFiles.length > 0 && (
                <>
                  <div style={{
                    padding: '8px 16px',
                    backgroundColor: '#f8f8f8',
                    borderBottom: '1px solid #e1dfdd'
                  }}>
                    <Caption1 style={{ color: '#605e5c', fontWeight: '600' }}>
                      {getLocalizedString('ExistingFilesIn', 'Existing Files in')} {getCombinedFolderPath() || getLocalizedString('Container', 'Container')}
                    </Caption1>
                  </div>
                  {existingFiles.map((existingFile, index) => (
                    <div key={`existing-${existingFile.name}-${index}`} style={{
                      display: 'flex',
                      alignItems: 'center',
                      padding: '12px 16px',
                      borderBottom: '1px solid #f3f2f1',
                      backgroundColor: '#f9f9f9'
                    }}>
                      {/* File Info */}
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: 'flex', alignItems: 'center', marginBottom: '4px' }}>
                          <Text 
                            style={{ 
                              fontWeight: '600', 
                              overflow: 'hidden',
                              textOverflow: 'ellipsis',
                              whiteSpace: 'nowrap',
                              marginRight: '8px'
                            }}
                          >
                            {existingFile.name}
                          </Text>
                          <CheckmarkCircleFilled style={{ color: '#0078d4', fontSize: '16px' }} />
                        </div>
                        <Caption1 style={{ color: '#605e5c' }}>
                          {formatFileSize(existingFile.size)} • {getLocalizedString('Uploaded', 'Uploaded')} {existingFile.lastModified.toLocaleDateString()}
                        </Caption1>
                      </div>
                      
                      {/* Action buttons */}
                      <div style={{ display: 'flex', gap: '8px' }}>
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={viewingFiles[existingFile.name] ? <Spinner size="tiny" /> : <EyeRegular />}
                          onClick={() => handleViewExistingFile(existingFile.name)}
                          disabled={viewingFiles[existingFile.name]}
                        >
                          {viewingFiles[existingFile.name] ? getLocalizedString('Generating', 'Downloading...') : getLocalizedString('View', 'Download')}
                        </Button>
                        {buttonDisplayMode === '0' && (
                          <>
                            {!showDeleteConfirmation[existingFile.name] ? (
                              <Button
                                appearance="subtle"
                                size="small"
                                icon={<DeleteRegular />}
                                onClick={() => showDeleteConfirmationButtons(existingFile.name)}
                                disabled={deletingFiles[existingFile.name]}
                              >
                                {getLocalizedString('Delete', 'Delete')}
                              </Button>
                            ) : (
                              <>
                                <Button
                                  appearance="subtle"
                                  size="small"
                                  icon={deletingFiles[existingFile.name] ? <Spinner size="tiny" /> : <CheckmarkRegular />}
                                  onClick={() => confirmDelete(existingFile.name)}
                                  disabled={deletingFiles[existingFile.name]}
                                  style={{ color: '#107c10' }}
                                >
                                  {deletingFiles[existingFile.name] ? getLocalizedString('Deleting', 'Deleting...') : getLocalizedString('Yes', 'Yes')}
                                </Button>
                                <Button
                                  appearance="subtle"
                                  size="small"
                                  icon={<DismissRegular />}
                                  onClick={() => cancelDeleteConfirmation(existingFile.name)}
                                  disabled={deletingFiles[existingFile.name]}
                                  style={{ color: '#d13438' }}
                                >
                                  {getLocalizedString('No', 'No')}
                                </Button>
                              </>
                            )}
                          </>
                        )}
                      </div>
                    </div>
                  ))}
                </>
              )}
              
              {/* New Files Section */}
              {fileStates.length > 0 && (
                <>
                  {existingFiles.length > 0 && (
                    <div style={{
                      padding: '8px 16px',
                      backgroundColor: '#f8f8f8',
                      borderBottom: '1px solid #e1dfdd'
                    }}>
                      <Caption1 style={{ color: '#605e5c', fontWeight: '600' }}>
                        {getLocalizedString('NewFiles', 'New Files')}
                      </Caption1>
                    </div>
                  )}
                  {fileStates.map((fileState, index) => (
                    <div key={`new-${fileState.file.name}-${index}`} style={{
                      display: 'flex',
                      alignItems: 'center',
                      padding: '12px 16px',
                      borderBottom: index < fileStates.length - 1 || existingFiles.length > 0 ? '1px solid #f3f2f1' : 'none',
                      backgroundColor: fileState.status === 'invalid' ? '#fef7f7' : 'transparent'
                    }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: 'flex', alignItems: 'center', marginBottom: '4px' }}>
                          <Text 
                            style={{ 
                              fontWeight: '600', 
                              overflow: 'hidden',
                              textOverflow: 'ellipsis',
                              whiteSpace: 'nowrap',
                              marginRight: '8px'
                            }}
                          >
                            {fileState.file.name}
                          </Text>
                          {fileState.status === 'uploading' && (
                            <Spinner size="tiny" style={{ marginRight: '8px' }} />
                          )}
                          {fileState.status === 'completed' && (
                            <CheckmarkCircleFilled style={{ color: '#107c10', fontSize: '16px' }} />
                          )}
                          {fileState.status === 'failed' && (
                            <ErrorCircleFilled style={{ color: '#d13438', fontSize: '16px' }} />
                          )}
                          {fileState.status === 'invalid' && (
                            <ErrorCircleFilled style={{ color: '#d13438', fontSize: '16px' }} />
                          )}
                        </div>
                        <Caption1 style={{ color: fileState.status === 'invalid' ? '#d13438' : '#605e5c' }}>
                          {formatFileSize(fileState.file.size)}
                          {fileState.status === 'completed' && ` • ${getLocalizedString('FileUploadedSuccessfully', 'File uploaded successfully')}`}
                          {fileState.status === 'failed' && ` • ${fileState.error || getLocalizedString('UploadFailed', 'Upload failed')}`}
                          {fileState.status === 'pending' && ` • ${getLocalizedString('ReadyToUpload', 'Ready to upload')}`}
                          {fileState.status === 'uploading' && ` • ${getLocalizedString('Uploading', 'Uploading...')}`}
                          {fileState.status === 'invalid' && ` • ${fileState.error || getLocalizedString('InvalidFileType', 'Invalid file type')}`}
                        </Caption1>
                        
                        {/* Progress Bar */}
                        {fileState.status === 'uploading' && (
                          <div style={{ marginTop: '8px' }}>
                            <ProgressBar 
                              value={fileState.progress / 100}
                              color="brand"
                            />
                            <Caption1 style={{ color: '#605e5c', marginTop: '2px' }}>
                              {fileState.progress}%
                            </Caption1>
                          </div>
                        )}
                        {fileState.status === 'completed' && (
                          <div style={{ marginTop: '8px' }}>
                            <ProgressBar 
                              value={1}
                              color="success"
                            />
                            <Caption1 style={{ color: '#107c10', marginTop: '2px' }}>
                              100% - {getLocalizedString('Complete', 'Complete')}
                            </Caption1>
                          </div>
                        )}
                      </div>

                      {/* Remove Button for pending/invalid files */}
                      {(fileState.status === 'pending' || fileState.status === 'invalid') && buttonDisplayMode === '0' && (
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={<DismissRegular />}
                          onClick={() => removeFile(fileState.file.name)}
                          style={{ marginLeft: '8px' }}
                        />
                      )}
                    </div>
                  ))}
                </>
              )}
                </>
              )}
            </div>
          </div>
        )}

        {/* Hidden file input */}
        <input
          ref={importFileRef}
          multiple={buttonAllowMultipleFiles}
          type="file"
          accept={buttonDisplayMode === '3' ? '*' : buttonAllowedFileTypes}
          onChange={onFileChange}
          style={{ display: "none" }}
        />
        </>
        )}
      </div>
    </FluentProvider>
  );
};
