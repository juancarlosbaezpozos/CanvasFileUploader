/**
 * Button utility functions and enums for Fluent UI component configuration
 *
 * This module provides a comprehensive set of utilities for configuring Fluent UI Button
 * components with type-safe enum conversions and style calculations. It handles the
 * conversion between string-based enum values (typically from configuration or props)
 * and the appropriate Fluent UI component properties.
 */

import { Button, type ButtonProps } from "@fluentui/react-components";

/**
 * Enumeration of button loading states for providing UI feedback during operations
 * These states help manage button appearance and behavior during async operations
 */
export enum ButtonLoadingStateEnum {
  /** Initial state before any operation starts */
  Initial = "initial",
  /** State during ongoing operation (typically shows spinner/disabled state) */
  Loading = "loading",
  /** State after operation completes successfully */
  Loaded = "loaded",
}

/**
 * Enumeration of button visual appearance styles matching Fluent UI design system
 * These values correspond directly to Fluent UI Button appearance prop values
 */
enum ButtonAppearanceEnum {
  /** Solid background with high contrast - used for primary actions */
  primary = 0,
  /** Outlined style with transparent background - used for secondary actions */
  secondary = 1,
  /** Outlined border with transparent background - alternative secondary style */
  outline = 2,
  /** Subtle background on hover - used for tertiary actions */
  subtle = 3,
  /** Completely transparent until hover - used for minimal UI impact */
  transparent = 4,
}

/**
 * Enumeration of button corner shapes for different design aesthetics
 * Controls the border-radius styling of button components
 */
enum ButtonShapeEnum {
  /** Standard rounded corners following design system guidelines */
  rounded = 0,
  /** Fully circular shape - typically used for icon-only buttons */
  circular = 1,
  /** Sharp square corners with no border-radius */
  square = 2,
}

/**
 * Enumeration of icon position relative to button text content
 * Determines the layout direction of icon and text within the button
 */
enum ButtonIconPositionEnum {
  /** Icon appears before (left of) the text content */
  before = 0,
  /** Icon appears after (right of) the text content */
  after = 1,
}

/**
 * Enumeration of button size presets following Fluent UI sizing standards
 * Each size affects padding, font-size, and overall button dimensions
 */
enum ButtonSizeEnum {
  /** Compact size for dense UI layouts */
  small = 0,
  /** Standard size for most use cases */
  medium = 1,
  /** Larger size for prominent actions or accessibility */
  large = 2,
}

/**
 * Enumeration of button alignment options for layout positioning
 * Controls how the button content is justified within its container
 */
enum ButtonAlign {
  /** Align content to the left (flex-start) */
  left = 0,
  /** Center align content horizontally */
  center = 1,
  /** Align content to the right (flex-end) */
  right = 2,
  /** Stretch content to fill available space */
  justify = 3,
}

/**
 * Converts button alignment enum value to corresponding CSS justify-content property
 *
 * This function safely parses string-based enum values and maps them to appropriate
 * CSS flexbox justify-content values for button content alignment.
 *
 * @param align - String representation of ButtonAlign enum value (0-3)
 * @returns CSS justify-content value for flexbox layout
 * @throws Error if the provided align value is invalid or out of range
 *
 * @example
 * ```typescript
 * const alignment = getButtonAlign("1"); // Returns "center"
 * const alignment = getButtonAlign("0"); // Returns "flex-start"
 * ```
 */
const getButtonAlign = (align: string): string => {
  const parsedEnum = Number.parseInt(align);
  if (isNaN(parsedEnum) || !(parsedEnum in ButtonAlign)) {
    throw new Error(`Invalid button align: ${align}`);
  }

  switch (ButtonAlign[parsedEnum]) {
    case ButtonAlign[ButtonAlign.left]:
      return "flex-start";
    case ButtonAlign[ButtonAlign.center]:
      return "center";
    case ButtonAlign[ButtonAlign.right]:
      return "flex-end";
    case ButtonAlign[ButtonAlign.justify]:
      return "stretch";
    default:
      return "center";
  }
};

/**
 * Enumeration of font weight options for button text styling
 * Maps to standard CSS font-weight values for consistent typography
 */
enum ButtonFontWeight {
  /** Bold font weight (700) for emphasis */
  bold = 0,
  /** Lighter font weight (300) for subtle text */
  lighter = 1,
  /** Normal font weight (400) - default weight */
  normal = 2,
  /** Semi-bold font weight (600) for medium emphasis */
  semibold = 3,
}

/**
 * Converts button font weight enum value to corresponding CSS font-weight property
 *
 * This function safely parses string-based enum values and maps them to standard
 * CSS font-weight values for consistent button text styling.
 *
 * @param weight - String representation of ButtonFontWeight enum value (0-3)
 * @returns CSS font-weight value as string
 * @throws Error if the provided weight value is invalid or out of range
 *
 * @example
 * ```typescript
 * const weight = getButtonFontWeight("0"); // Returns "bold"
 * const weight = getButtonFontWeight("2"); // Returns "normal"
 * ```
 */
const getButtonFontWeight = (weight: string): string => {
  const parsedEnum = Number.parseInt(weight);
  if (isNaN(parsedEnum) || !(parsedEnum in ButtonFontWeight)) {
    throw new Error(`Invalid button font weight: ${weight}`);
  }

  switch (ButtonFontWeight[parsedEnum]) {
    case ButtonFontWeight[ButtonFontWeight.bold]:
      return "bold";
    case ButtonFontWeight[ButtonFontWeight.semibold]:
      return "semibold";
    case ButtonFontWeight[ButtonFontWeight.normal]:
      return "normal";
    case ButtonFontWeight[ButtonFontWeight.lighter]:
      return "lighter";
    default:
      return "normal";
  }
};

/**
 * Converts button appearance enum value to Fluent UI Button appearance prop
 *
 * This function provides type-safe conversion from string-based configuration values
 * to Fluent UI Button component appearance properties, ensuring design consistency.
 *
 * @param appearance - String representation of ButtonAppearanceEnum value (0-4)
 * @returns Fluent UI ButtonProps appearance value
 * @throws Error if the provided appearance value is invalid or out of range
 *
 * @example
 * ```typescript
 * const appearance = getButtonAppearance("0"); // Returns "primary"
 * const appearance = getButtonAppearance("2"); // Returns "outline"
 * ```
 */
export const getButtonAppearance = (
  appearance: string
): ButtonProps["appearance"] => {
  const parsedEnum = Number.parseInt(appearance);
  if (isNaN(parsedEnum) || !(parsedEnum in ButtonAppearanceEnum)) {
    throw new Error(`Invalid button appearance: ${appearance}`);
  }
  return ButtonAppearanceEnum[parsedEnum] as ButtonProps["appearance"];
};

/**
 * Converts icon position enum value to Fluent UI Button iconPosition prop
 *
 * This function handles the conversion of string-based configuration values to
 * Fluent UI Button iconPosition properties for proper icon-text layout.
 *
 * @param iconPosition - String representation of ButtonIconPositionEnum value (0-1)
 * @returns Fluent UI ButtonProps iconPosition value ("before" or "after")
 * @throws Error if the provided iconPosition value is invalid or out of range
 *
 * @example
 * ```typescript
 * const position = getButtonIconPosition("0"); // Returns "before"
 * const position = getButtonIconPosition("1"); // Returns "after"
 * ```
 */
export const getButtonIconPosition = (
  iconPosition: string
): ButtonProps["iconPosition"] => {
  const parsedEnum = Number.parseInt(iconPosition);
  if (isNaN(parsedEnum) || !(parsedEnum in ButtonIconPositionEnum)) {
    throw new Error(`Invalid button icon position: ${iconPosition}`);
  }
  return ButtonIconPositionEnum[parsedEnum] as ButtonProps["iconPosition"];
};

/**
 * Converts shape enum value to Fluent UI Button shape prop
 *
 * This function provides type-safe conversion from configuration values to
 * Fluent UI Button shape properties for consistent visual styling.
 *
 * @param shape - String representation of ButtonShapeEnum value (0-2)
 * @returns Fluent UI ButtonProps shape value ("rounded", "circular", or "square")
 * @throws Error if the provided shape value is invalid or out of range
 *
 * @example
 * ```typescript
 * const shape = getButtonShape("0"); // Returns "rounded"
 * const shape = getButtonShape("1"); // Returns "circular"
 * ```
 */
export const getButtonShape = (shape: string): ButtonProps["shape"] => {
  const parsedEnum = Number.parseInt(shape);
  if (isNaN(parsedEnum) || !(parsedEnum in ButtonShapeEnum)) {
    throw new Error(`Invalid button shape: ${shape}`);
  }
  return ButtonShapeEnum[parsedEnum] as ButtonProps["shape"];
};

/**
 * Converts size enum value to Fluent UI Button size prop
 *
 * This function handles the conversion from string-based configuration values to
 * Fluent UI Button size properties for consistent component sizing.
 *
 * @param size - String representation of ButtonSizeEnum value (0-2)
 * @returns Fluent UI ButtonProps size value ("small", "medium", or "large")
 * @throws Error if the provided size value is invalid or out of range
 *
 * @example
 * ```typescript
 * const size = getButtonSize("1"); // Returns "medium"
 * const size = getButtonSize("2"); // Returns "large"
 * ```
 */
export const getButtonSize = (size: string): ButtonProps["size"] => {
  const parsedEnum = Number.parseInt(size);
  if (isNaN(parsedEnum) || !(parsedEnum in ButtonSizeEnum)) {
    throw new Error(`Invalid button size: ${size}`);
  }
  return ButtonSizeEnum[parsedEnum] as ButtonProps["size"];
};

/**
 * Creates a comprehensive button style object with dimensions, alignment, and typography
 *
 * This function combines multiple style properties into a single object that can be
 * applied to Fluent UI Button components. It handles pixel-based sizing, content
 * alignment, and font weight styling in a type-safe manner.
 *
 * @param width - Button width in pixels
 * @param height - Button height in pixels
 * @param align - String representation of ButtonAlign enum value for content justification
 * @param weight - String representation of ButtonFontWeight enum value for text styling
 * @returns Complete style object compatible with Fluent UI ButtonProps["style"]
 *
 * @example
 * ```typescript
 * const buttonStyle = getButtonStyle(120, 40, "1", "0");
 * // Returns: { width: "120px", height: "40px", justifyContent: "center", fontWeight: "bold" }
 * ```
 */
export const getButtonStyle = (
  width: number,
  height: number,
  align: string,
  weight: string
): ButtonProps["style"] => {
  return {
    width: `${width}px`,
    height: `${height}px`,
    justifyContent: `${getButtonAlign(align)}`,
    fontWeight: `${getButtonFontWeight(weight)}`,
  };
};

/**
 * Determines button visibility based on current loading state and configured display mode
 *
 * This function implements the business logic for when buttons should be visible or hidden
 * based on the application's loading state and the configured display preferences.
 *
 * Display Mode Values:
 * - "0": Edit mode - Interactive button (visible except in loaded state)
 * - "1": Disabled mode - Button visible but disabled in all states
 * - "2": View mode - Button visible but disabled in all states
 * - "3": Admin mode - Configuration interface for runtime property changes
 * - Other values: Button hidden in loaded state
 *
 * @param loadingState - Current loading state of the application ("initial", "loading", "loaded")
 * @param buttonDisplayMode - String configuration value determining visibility rules
 * @returns Boolean indicating whether the button should be displayed
 *
 * @example
 * ```typescript
 * const shouldShow = getDisplayMode("loading", "1"); // Returns true
 * const shouldShow = getDisplayMode("loaded", "1");   // Returns true
 * const shouldShow = getDisplayMode("loaded", "0");   // Returns false
 * const shouldShow = getDisplayMode("initial", "3");  // Returns true (admin mode)
 * ```
 *
 * @todo Consider refactoring to use enum values for buttonDisplayMode for better type safety
 */
export const getDisplayMode = (
  loadingState: string,
  buttonDisplayMode: string
): boolean => {
  // Admin mode (3) is always visible for configuration
  if (buttonDisplayMode === "3") {
    return true;
  }

  // Always show button during loading state regardless of display mode
  return loadingState === "loading"
    ? true
    : loadingState === "loaded"
    ? buttonDisplayMode === "1" || buttonDisplayMode === "2"
    : buttonDisplayMode === "1" || buttonDisplayMode === "2";
};
