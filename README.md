# Files Import Control Documentation

## Purpose of the Control

The **Files Import Control** (Capgemini Fluent Files Import Control) is a PowerApps Component Framework (PCF) control designed to provide a comprehensive file upload and management solution for Microsoft Power Platform applications. The control serves the following primary purposes:

- **File Upload Management**: Enables users to select and upload single or multiple files with progress tracking and validation
- **Azure Blob Storage Integration**: Seamlessly integrates with Azure Blob Storage through Power Automate Cloud Flows for scalable file storage
- **File Operations**: Provides complete file lifecycle management including upload, view, list, and delete operations
- **Business Application Integration**: Designed specifically for business applications requiring file attachment capabilities tied to records
- **Compliance and Validation**: Enforces file type restrictions and size limits to ensure compliance with organizational policies
- **Dual Operation Modes**: Supports both Cloud Flow-based operations (for production environments) and JSON-based operations (for development/testing)

The control is particularly suited for scenarios where organizations need to associate files with specific business records while maintaining centralized storage in Azure and enforcing consistent file management policies.

## Properties Data Dictionary

### General Properties

| Property | Type | Required | Default | Description |
|----------|------|----------|---------|-------------|
| **Text** | SingleLine.Text | Yes | "File upload" | Text displayed on the upload button |
| **ButtonIcon** | SingleLine.Text | Yes | "Attach" | Icon name to display in the button (Fluent UI icon names) |
| **IconStyle** | Enum | Yes | Outline (0) | Visual style of the icon - Outline (0) or Filled (1) |
| **Visible** | TwoOptions | Yes | true | Controls button visibility |

### Behavior Properties

| Property | Type | Required | Default | Description |
|----------|------|----------|---------|-------------|
| **DisplayMode** | Enum | Yes | Edit (0) | Controls the operational mode: Edit (0), Disabled (1), View (2), Admin (3) |
| **AllowMultipleFiles** | TwoOptions | No | true | Enables selection of multiple files simultaneously |
| **AllowedFileTypes** | SingleLine.Text | No | "*" | Semicolon or comma-separated list of allowed file extensions (e.g., ".pdf;.docx;.txt") |
| **AllowDropFiles** | TwoOptions | No | true | Enables drag-and-drop file selection |
| **AllowDropFilesText** | SingleLine.Text | No | "Drop files here..." | Text displayed during drag-and-drop operations |
| **MaxTotalFileSizeMB** | Whole.None | No | 20 | Maximum combined size of all files in megabytes |

### Visual Appearance Properties

| Property | Type | Required | Default | Description |
|----------|------|----------|---------|-------------|
| **Width** | Whole.None | Yes | 128 | Button width in pixels |
| **Height** | Whole.None | Yes | 32 | Button height in pixels |
| **Appearance** | Enum | Yes | Primary (0) | Button style: Primary (0), Secondary (1), Outline (2), Subtle (3), Transparent (4) |
| **Align** | Enum | Yes | Center (1) | Button text alignment: Left (0), Center (1), Right (2), Justify (3) |
| **FontWeight** | Enum | Yes | Normal (2) | Button text weight: Bold (0), Lighter (1), Normal (2), Semibold (3) |
| **IconPosition** | Enum | Yes | Before (0) | Icon placement: Before (0) or After (1) text |
| **Shape** | Enum | Yes | Rounded (0) | Button shape: Rounded (0), Circular (1), Square (2) |
| **ButtonSize** | Enum | Yes | Small (0) | Button size: Small (0), Medium (1), Large (2) |
| **DisabledFocusable** | TwoOptions | No | false | Allows focus on disabled buttons for accessibility |
| **ShowActionSpinner** | TwoOptions | No | true | Shows spinner animation during upload operations |

### Cloud Flow Configuration Properties

| Property | Type | Required | Default | Description |
|----------|------|----------|---------|-------------|
| **CloudFlowUploadUrl** | SingleLine.Text | No | null | Complete trigger URL for the file upload Power Automate flow |
| **CloudFlowListFilesUrl** | SingleLine.Text | No | null | Complete trigger URL for the list files Power Automate flow |
| **CloudFlowDeleteUrl** | SingleLine.Text | No | null | Complete trigger URL for the delete file Power Automate flow |
| **CloudFlowGenerateViewUrl** | SingleLine.Text | No | null | Complete trigger URL for generating SAS URLs to view files |
| **ContainerPath** | SingleLine.Text | No | null | Azure Storage account name or container path |
| **ListFilesFolderName** | SingleLine.Text | No | null | Folder name for organizing and listing files |
| **RecordUid** | SingleLine.Text | No | null | Unique identifier linking files to specific records. When empty, control operates in JSON mode |

### Output Properties

| Property | Type | Description |
|----------|------|-------------|
| **FilesAsJSON** | SingleLine.Text | JSON string containing selected/uploaded file information |
| **ExistingFiles** | Multiple | Collection of existing files in the specified folder |
| **UploadResults** | Multiple | Results from Cloud Flow upload operations |
| **LastUploadStatus** | Enum | Status of the last upload: None (0), InProgress (1), Completed (2), Failed (3) |

## Edit Mode (DisplayMode = 0)

### Functional Characteristics

**Interactive File Operations:**
- Full file selection capabilities through button click or drag-and-drop
- Real-time file validation against allowed types and size limits
- Progress tracking during upload operations with visual feedback
- Multiple file selection and batch upload support
- File removal from upload queue before processing

**User Interface Features:**
- Responsive upload button with customizable appearance
- Drag-and-drop zone with visual feedback during hover
- Progress bars showing individual file upload status
- File list display with validation status indicators
- Error messaging for validation failures

**Cloud Flow Integration:**
- When `RecordUid` is provided: Files are uploaded to Azure Blob Storage via Cloud Flows
- File operations include upload, list existing files, view files (SAS URL generation), and delete
- Automatic folder organization using `ListFilesFolderName` and `RecordUid`

**JSON Mode Operation:**
- When `RecordUid` is empty: Files are processed as Base64-encoded JSON data
- Suitable for development environments or scenarios without Cloud Flow integration
- Files are stored in component state and output as JSON

**Validation and Compliance:**
- File type validation against `AllowedFileTypes` property
- Total file size validation against `MaxTotalFileSizeMB` limit
- Real-time feedback on validation status

### Limitations in Edit Mode

- Depends on proper Cloud Flow configuration for production file storage
- File size limitations imposed by Power Automate and browser capabilities
- Network connectivity required for Cloud Flow operations

## Display Mode (DisplayMode = 1 & 2)

### Functional Characteristics

**View-Only Operations:**
- Button is disabled but remains visible for context
- No file selection or upload capabilities
- Existing files are displayed in read-only mode
- File viewing through SAS URL generation (if configured)

**File Listing:**
- Displays existing files in the configured folder path
- Shows file metadata (name, size, last modified date)
- Provides view/download links for existing files
- Loading states during file list retrieval

**Status Information:**
- Shows configured folder path and container information
- Displays file count and total size information
- Maintains visual consistency with edit mode

### Limitations in Display/View Mode

- No file upload or modification capabilities
- Cannot delete existing files
- No drag-and-drop functionality
- Cannot change file validation settings
- Limited to viewing operations only

## Administration Mode (DisplayMode = 3)

### Functional Characteristics

**Runtime Configuration Interface:**
- Dedicated admin panel for modifying control behavior
- Real-time property updates without design-time changes
- Persistent configuration storage in browser localStorage
- Override capabilities for design-time property values

**Configurable Properties:**
- **Maximum Total File Size**: Runtime adjustment of file size limits (1-100 MB)
- **Allowed File Types**: Dynamic modification of permitted file extensions
- Configuration validation and feedback

**Administrative Features:**
- Configuration status display showing current vs. original settings
- Reset functionality to restore original design-time values
- Apply button for confirming changes (though changes are applied automatically)
- Visual indicators for configuration state

**Persistence Mechanism:**
- Settings stored in browser localStorage with key `fileUploadControlAdminConfig`
- Automatic loading of saved configuration on control initialization
- Fallback to design-time properties if saved configuration is invalid

### Limitations in Administration Mode

- Cannot modify Cloud Flow URLs or container paths
- Limited to file validation properties only
- Configuration is browser-specific (localStorage-based)
- Requires manual display mode change to exit admin interface
- No export/import functionality for configurations across environments

## Known Dependencies

### Platform Dependencies

- **Microsoft PowerApps Component Framework (PCF)**: Core platform requirement
- **Power Platform**: Host environment for the control
- **Microsoft Power Automate**: Required for Cloud Flow-based file operations
- **Azure Blob Storage**: Storage backend when using Cloud Flow integration

### JavaScript/TypeScript Dependencies

- **React 16.14.0**: Core UI framework
- **@fluentui/react-components 9.46.2**: Microsoft Fluent UI component library
- **TypeScript 4.8.4+**: Language and type checking
- **tinycolor2**: Color manipulation utilities

### Development Dependencies

- **pcf-scripts**: PowerApps Component Framework build tools
- **ESLint**: Code linting and quality
- **@types/powerapps-component-framework**: TypeScript definitions for PCF

### External Service Dependencies

- **logic.azure.com**: Power Automate cloud flows execution domain
- **Azure Storage APIs**: Indirect dependency through Power Automate flows
- **Microsoft authentication services**: For Cloud Flow authentication

### Browser Requirements

- **Modern browsers supporting ES6+**: Chrome, Edge, Firefox, Safari
- **localStorage support**: For admin configuration persistence
- **File API support**: For file reading and drag-and-drop operations
- **Fetch API support**: For Cloud Flow HTTP requests

### Power Automate Flow Requirements

The control expects specific Power Automate flows with standardized request/response schemas:

1. **Upload Flow**: Accepts Base64 file content, returns success status and file URL
2. **List Files Flow**: Returns array of file metadata for specified folder
3. **Delete Flow**: Accepts file identifier, returns deletion status
4. **Generate View URL Flow**: Returns time-limited SAS URLs for file access

### Localization Dependencies

- **Resource files (.resx)**: Support for multiple languages (1033-English, 1034-Spanish, 1046-Portuguese)
- **Power Platform localization services**: For runtime string localization

### Security Dependencies

- **Same-origin policy compliance**: For localStorage and fetch operations
- **Content Security Policy (CSP)**: Must allow inline styles and Power Automate domains
- **CORS configuration**: Azure Storage and Power Automate endpoints must allow cross-origin requests from Power Platform domains
