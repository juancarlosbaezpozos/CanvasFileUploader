/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    Text: ComponentFramework.PropertyTypes.StringProperty;
    ButtonIcon: ComponentFramework.PropertyTypes.StringProperty;
    IconStyle: ComponentFramework.PropertyTypes.EnumProperty<"0" | "1">;
    Visible: ComponentFramework.PropertyTypes.TwoOptionsProperty;
    DisplayMode: ComponentFramework.PropertyTypes.EnumProperty<"0" | "1" | "2" | "3">;
    Width: ComponentFramework.PropertyTypes.WholeNumberProperty;
    Height: ComponentFramework.PropertyTypes.WholeNumberProperty;
    Appearance: ComponentFramework.PropertyTypes.EnumProperty<"0" | "1" | "2" | "3" | "4">;
    Align: ComponentFramework.PropertyTypes.EnumProperty<"0" | "1" | "2" | "3">;
    FontWeight: ComponentFramework.PropertyTypes.EnumProperty<"0" | "1" | "2" | "3">;
    IconPosition: ComponentFramework.PropertyTypes.EnumProperty<"0" | "1">;
    Shape: ComponentFramework.PropertyTypes.EnumProperty<"0" | "1" | "2">;
    ButtonSize: ComponentFramework.PropertyTypes.EnumProperty<"0" | "1" | "2">;
    DisabledFocusable: ComponentFramework.PropertyTypes.TwoOptionsProperty;
    AllowMultipleFiles: ComponentFramework.PropertyTypes.TwoOptionsProperty;
    AllowedFileTypes: ComponentFramework.PropertyTypes.StringProperty;
    AllowDropFiles: ComponentFramework.PropertyTypes.TwoOptionsProperty;
    AllowDropFilesText: ComponentFramework.PropertyTypes.StringProperty;
    ShowActionSpinner: ComponentFramework.PropertyTypes.TwoOptionsProperty;
    MaxTotalFileSizeMB: ComponentFramework.PropertyTypes.WholeNumberProperty;
    CloudFlowUploadUrl: ComponentFramework.PropertyTypes.StringProperty;
    CloudFlowListFilesUrl: ComponentFramework.PropertyTypes.StringProperty;
    CloudFlowDeleteUrl: ComponentFramework.PropertyTypes.StringProperty;
    CloudFlowDownloadUrl: ComponentFramework.PropertyTypes.StringProperty;
    ContainerPath: ComponentFramework.PropertyTypes.StringProperty;
    ListFilesFolderName: ComponentFramework.PropertyTypes.StringProperty;
    RecordUid: ComponentFramework.PropertyTypes.StringProperty;
}
export interface IOutputs {
    Text?: string;
    ButtonIcon?: string;
    AllowedFileTypes?: string;
    AllowDropFilesText?: string;
    FilesAsJSON?: string;
    ExistingFiles?: string;
    UploadResults?: string;
    LastUploadStatus?: string;
}
