# CLAUDE.md - AI Assistant Guidelines

This file provides guidance for AI assistants working with this repository.

## Project Overview

This is a new repository that is being set up. Update this section with:
- Project name and purpose
- Core functionality and features
- Target users/audience

## Repository Structure

```
/home/user/dev/
├── CLAUDE.md          # AI assistant guidelines (this file)
└── .git/              # Git repository
```

> **Note**: This is a newly initialized repository. Update this structure diagram as the project grows.

## Technology Stack

Document the technologies used in this project:

- **Language**: (e.g., TypeScript, Python, Go)
- **Framework**: (e.g., React, Express, Django)
- **Database**: (e.g., PostgreSQL, MongoDB)
- **Testing**: (e.g., Jest, pytest)
- **Build Tools**: (e.g., Vite, webpack, esbuild)

## Development Setup

### Prerequisites

List required tools and versions:
- Node.js / Python / etc.
- Package manager (npm, yarn, pnpm, pip)
- Any other dependencies

### Installation

```bash
# Clone the repository
git clone <repository-url>
cd dev

# Install dependencies
# npm install / pip install -r requirements.txt / etc.
```

## Common Commands

| Command | Description |
|---------|-------------|
| `npm install` | Install dependencies |
| `npm run dev` | Start development server |
| `npm run build` | Build for production |
| `npm run test` | Run tests |
| `npm run lint` | Run linter |
| `npm run format` | Format code |

> Update these commands based on the actual project setup.

## Code Conventions

### File Naming
- Use kebab-case for file names: `my-component.ts`
- Use PascalCase for React components: `MyComponent.tsx`
- Use camelCase for utility files: `myHelper.ts`

### Code Style
- Follow the project's linter configuration
- Use consistent indentation (spaces vs tabs)
- Add meaningful comments for complex logic
- Keep functions focused and single-purpose

### Commit Messages
- Use conventional commits format: `type(scope): description`
- Types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`
- Keep messages concise but descriptive

## Architecture Guidelines

### Directory Organization

Describe the intended directory structure:

```
src/
├── components/     # Reusable UI components
├── pages/          # Page-level components/routes
├── services/       # API and external service integrations
├── utils/          # Utility functions and helpers
├── types/          # TypeScript type definitions
├── hooks/          # Custom React hooks (if applicable)
└── constants/      # Application constants
```

### Design Patterns

Document key patterns used in this project:
- Component patterns
- State management approach
- Error handling strategies
- API integration patterns

## Testing Guidelines

### Test Structure
- Unit tests alongside source files or in `__tests__` directories
- Integration tests in dedicated `tests/` directory
- E2E tests in `e2e/` directory

### Running Tests

```bash
# Run all tests
npm test

# Run tests in watch mode
npm run test:watch

# Run with coverage
npm run test:coverage
```

## Environment Configuration

### Environment Variables

Document required environment variables:

| Variable | Description | Required |
|----------|-------------|----------|
| `DATABASE_URL` | Database connection string | Yes |
| `API_KEY` | External API key | Yes |
| `NODE_ENV` | Environment (development/production) | No |

### Configuration Files

- `.env` - Local environment variables (not committed)
- `.env.example` - Template for environment variables

## Git Workflow

### Branch Naming
- `feature/description` - New features
- `fix/description` - Bug fixes
- `docs/description` - Documentation updates
- `refactor/description` - Code refactoring

### Pull Request Process
1. Create a feature branch from main
2. Make changes and commit with clear messages
3. Push branch and create PR
4. Request review from team members
5. Address feedback and merge

## AI Assistant Notes

### When Working on This Codebase

1. **Read before modifying**: Always read existing code before making changes
2. **Follow existing patterns**: Match the coding style already in use
3. **Keep changes focused**: Only modify what's necessary for the task
4. **Test your changes**: Run tests after making modifications
5. **Check for security**: Avoid introducing vulnerabilities (XSS, injection, etc.)

### Common Tasks

#### Adding a New Feature
1. Understand requirements fully
2. Check for existing similar implementations
3. Follow established patterns
4. Add appropriate tests
5. Update documentation if needed

#### Fixing a Bug
1. Reproduce the issue first
2. Identify root cause
3. Implement minimal fix
4. Add regression test
5. Verify fix doesn't break other functionality

#### Refactoring Code
1. Ensure tests exist before refactoring
2. Make incremental changes
3. Run tests after each change
4. Keep commits atomic and reversible

### Files to Avoid Modifying

- Lock files (`package-lock.json`, `yarn.lock`) - only through package manager
- Generated files in `dist/`, `build/`, `.next/`
- IDE configuration files unless specifically requested

## Troubleshooting

### Common Issues

Document common problems and solutions:

| Issue | Solution |
|-------|----------|
| Dependencies not installing | Clear node_modules and reinstall |
| Tests failing | Check for missing env variables |
| Build errors | Verify TypeScript types are correct |

## Resources

- Project documentation: (link)
- API documentation: (link)
- Design system: (link)
- Team wiki: (link)

---

*Last updated: 2026-01-31*
*Update this file as the project evolves to keep AI assistants informed of current practices.*
