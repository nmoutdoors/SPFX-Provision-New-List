# SharePoint List Provisioning Web Part

## Summary

A SharePoint Framework (SPFx) web part that demonstrates automated list provisioning. This web part creates a "ProjectsNew" list with custom fields and sample data.

## Features

This web part demonstrates the following concepts:
- Automated SharePoint list creation
- Custom field provisioning (Status field with choice values)
- Sample data generation
- React-based UI using Fluent UI components
- SharePoint REST API integration
- User information retrieval

## Prerequisites

- Node.js version 18.17.1 or higher (but lower than 19.0.0)
- SharePoint Online tenant
- Appropriate permissions to create lists in your SharePoint site

## Getting Started

1. Clone this repository
2. Navigate to the project directory
3. Run the following commands:
   ```bash
   npm install
   gulp serve
   ```
4. When prompted, add the web part to your SharePoint page
5. Use the property pane to provision the new list

## Solution Details

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| provision   | [Your Name] ([Your Company])                            |

## Version History

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | [Current Date]   | Initial release |

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## References

- [SharePoint Framework Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [SharePoint REST API Reference](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
