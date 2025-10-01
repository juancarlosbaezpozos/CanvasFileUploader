/**
 * Fluent UI React Icons mapping utility
 * 
 * This module provides a centralized way to manage and access Fluent UI icons
 * throughout the application. It imports only the specific icons that are needed
 * to optimize bundle size and provides both component and JSX element access.
 */

// Import only the icons that are actually needed - this significantly reduces bundle size
// by avoiding importing the entire icon library
import {
  AttachArrowRightRegular,
  AttachRegular,
  AttachTextRegular,
  AttachArrowRightFilled,
  AttachFilled,
  AttachTextFilled,
  CheckmarkFilled,
  EyeRegular,
  EyeFilled,
  DeleteRegular,
  DeleteFilled,
} from "@fluentui/react-icons";
import * as React from "react";

/**
 * Retrieves the appropriate Fluent UI icon component based on name and style preference
 * 
 * This function returns the actual React component (not a rendered JSX element),
 * which is useful when you need to pass the component to other components or
 * when you want to render it with custom props.
 * 
 * @param iconName - The name of the icon (without Regular/Filled suffix)
 * @param isFilled - Whether to return the filled variant of the icon (default: false)
 * @returns React component for the requested icon, or AttachRegular as fallback
 * 
 * @example
 * ```tsx
 * const IconComponent = getIconComponent('Delete', true);
 * return <IconComponent className="my-icon" />;
 * ```
 */
export function getIconComponent(
  iconName: string,
  isFilled: boolean = false
): React.ComponentType<any> {
  // Map of icon names to their regular (outline) variants
  const regularIcons: Record<string, React.ComponentType<any>> = {
    AttachArrowRight: AttachArrowRightRegular,
    Attach: AttachRegular,
    AttachText: AttachTextRegular,
    Eye: EyeRegular,
    Delete: DeleteRegular,
  };

  // Map of icon names to their filled (solid) variants
  const filledIcons: Record<string, React.ComponentType<any>> = {
    AttachArrowRight: AttachArrowRightFilled,
    Attach: AttachFilled,
    AttachText: AttachTextFilled,
    Checkmark: CheckmarkFilled, // Note: Only available in filled variant
    Eye: EyeFilled,
    Delete: DeleteFilled,
  };

  // Select the appropriate mapping based on the requested style
  const mapping = isFilled ? filledIcons : regularIcons;
  
  // Return the requested icon component or fallback to default attach icon
  return mapping[iconName] || AttachRegular;
}

/**
 * Legacy function that returns a ready-to-use JSX element for the requested icon
 * 
 * This function is maintained for backward compatibility with existing code.
 * It internally uses getIconComponent() and renders the component as JSX.
 * For new code, consider using getIconComponent() directly for better flexibility.
 * 
 * @param iconName - The name of the icon (without Regular/Filled suffix)
 * @param isFilled - Whether to return the filled variant of the icon (default: false)
 * @returns JSX element ready for rendering
 * 
 * @example
 * ```tsx
 * return <div>{getIcon('Eye', false)}</div>;
 * ```
 * 
 * @deprecated Consider using getIconComponent() for new implementations
 */
export function getIcon(
  iconName: string,
  isFilled: boolean = false
): React.ReactNode {
  // Get the component using the primary function
  const IconComponent = getIconComponent(iconName, isFilled);
  
  // Return the rendered JSX element
  return <IconComponent />;
}
