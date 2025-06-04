# Claude Configuration

## IMPORTANT: Multiple CLAUDE.md Files Rule
- When updating this CLAUDE.md file, ALWAYS search for all copies across the project
- Update every instance found to keep them synchronized
- Use Glob tool with pattern "**/CLAUDE.md" to find all instances
- This ensures consistency across the entire codebase

## Code Style Preferences
- NO emojis in any code, documentation, or README files
- NO Claude references or attribution in code
- NO "Generated with Claude" messages
- Clean, professional documentation style
- Focus on technical content without decorative elements

## Repository Standards
- Use standard markdown formatting without emoji decorations
- Professional tone in all documentation
- Technical accuracy over visual appeal
- Clean commit messages without attribution footers

## PowerShell Specific Guidelines
- Use approved verbs for cmdlet names (Get-, Set-, New-, Remove-, etc.)
- Follow Pascal case for function names
- Use proper parameter attributes and validation
- Include comment-based help for all public functions
- Prefer pipelines over loops where appropriate
- Use Write-Verbose for debug output, not Write-Host

## Security Practices
- Never hardcode credentials or secrets
- Use SecureString for sensitive data
- Validate all user input
- Follow principle of least privilege
- Use certificate-based authentication where possible

## Project Structure
- Keep scripts modular and focused on single responsibilities
- Place utility functions in separate modules
- Use consistent file naming conventions
- Store certificates and keys in the certificates/ directory
- Keep Azure-specific code separate from general utilities

## Testing and Validation
- Always test scripts with -WhatIf when applicable
- Include error handling with try/catch blocks
- Use proper exit codes
- Validate parameters before processing
- Test with different PowerShell versions if compatibility is required

## Documentation Requirements
- Include purpose and usage examples in script headers
- Document all parameters with clear descriptions
- Provide examples for complex operations
- Keep README files updated with current functionality
- Document any external dependencies

## Git Workflow
- Use descriptive commit messages
- Keep commits focused on single changes
- Test locally before committing
- Never commit sensitive data or credentials
- Use .gitignore for generated files and secrets

## PowerShell Honeypot Script Management - CRITICAL
- **ALWAYS** automatically commit new scripts and major updates to the pwsh-honeypot repository
- **IMMEDIATELY** after creating or significantly updating any PowerShell script in this repo, run:
  ```bash
  git add [script-path]
  git commit -m "[brief description of script/change]"
  ```
- This ensures the remote repository stays current with all script development
- Exception: Only skip auto-commit for experimental/debugging code that isn't ready for production
- Always provide clear, professional commit messages that describe the script's purpose
- After committing, remind user to push changes: "Please push to GitHub using: git push origin main"

## Website Learning Section Updates
- The "What I Learned Building This" section in `site/index.html` is where we track project learnings and challenges
- When user mentions adding learnings or challenges from the cloud resume project, update the HTML learning section (NOT this CLAUDE.md file)
- Add new learning items after the most recent ones but before the "Collapsible older items section" comment
- Each learning item should follow the existing format with Challenge/Solution structure
- Keep descriptions concise and focused on the key technical insight

## Git Push Rules - CRITICAL REMINDERS
- **NEVER** attempt to push to GitHub directly due to SSH authentication requirements
- **ALWAYS** remind the user to push after making commits - this is mandatory, not optional
- After any git commit operation, **IMMEDIATELY** tell the user: "Please push the changes to GitHub in a separate terminal window using: `git push origin <branch-name>`"
- **NEVER** recommend or suggest HTTPS authentication for Git operations - GitHub has deprecated password authentication for HTTPS
- When work sessions involve git commits, end the session by reminding about pending pushes
- If multiple commits are made during a session, remind about pushing at the end of each logical work unit

## Terminal State Management
- Always ensure bracket paste mode is disabled at session end
- If terminal state is modified during work, provide reset commands
- Before session ends, run: printf '\e[?2004l' to disable bracket pasting
- Remind user to reset terminal if any persistent state changes occur
- When demonstrating commands that might affect terminal state, include cleanup steps