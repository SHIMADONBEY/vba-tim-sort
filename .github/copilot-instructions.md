# Copilot Instructions
This file contains instructions for GitHub Copilot. It uses a YAML front matter header to specify which files the instructions apply to and a description of the instructions. The instructions themselves are written in markdown format.

## Basic Usage

- Comment English instructions in the code to guide Copilot's suggestions. If commited by Japanese speakers, the instructions may be in Japanese.
- Use clear and concise language to describe the desired code behavior or structure.
- Avoid ambiguous or overly complex instructions that may confuse Copilot.
- Regularly review and update the instructions as needed to ensure they remain relevant and effective.

## Visual Basic for Applications (VBA) Specific Instructions

- When writing instructions for VBA code, consider the unique syntax and conventions of VBA.
- Provide context about the purpose of the code and any specific requirements or constraints.
- If applicable, include examples of expected input and output to clarify the desired behavior.
- Be mindful of the limitations of Copilot when working with VBA, as it may not always generate code that is compatible with VBA's specific features and quirks.
- Use `Option Explicit` in your VBA code to help catch undeclared variables and improve code quality, and mention this in the instructions if relevant.
- Don't use `On Error Resume Next` in VBA code, as it can lead to unhandled errors and make debugging difficult. Instead, use structured error handling with `On Error GoTo` and mention this in the instructions if relevant.
- Avoid using `Stop` statements in VBA code, as they can halt execution and make it difficult to debug. Instead, use proper error handling and logging, and mention this in the instructions if relevant.
- Avoid using `Debug.Print` statements in VBA code, as they can clutter the output and make it difficult to read. Instead, use proper logging mechanisms and mention this in the instructions if relevant.
- Do not include trailing whitespace in your VBA code, as it can lead to formatting issues and make the code harder to read. Ensure that each line ends with a single newline character (LF) and mention this in the instructions if relevant.
- Do not use windows api calls in VBA code, as they can lead to compatibility issues and make the code less portable. Instead, use built-in VBA functions and features, and mention this in the instructions if relevant.
