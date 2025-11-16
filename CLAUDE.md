# CLAUDE.md - AI Assistant Guide for HealthTracker

> **Last Updated:** 2025-11-16
> **Repository:** JuMaD/HealthTracker
> **License:** MIT
> **Status:** Initial Setup Phase

## Table of Contents

- [Project Overview](#project-overview)
- [Current Repository State](#current-repository-state)
- [Recommended Architecture](#recommended-architecture)
- [Development Workflows](#development-workflows)
- [Code Conventions](#code-conventions)
- [AI Assistant Guidelines](#ai-assistant-guidelines)
- [Security Considerations](#security-considerations)
- [Testing Strategy](#testing-strategy)

---

## Project Overview

**HealthTracker** is a health and fitness tracking application designed to help users monitor various health metrics, activities, and wellness data.

### Purpose
- Track health metrics (weight, blood pressure, heart rate, etc.)
- Monitor fitness activities and exercise routines
- Manage dietary information and nutrition
- Visualize health trends over time
- Set and track health goals

### Target Users
- Individuals managing personal health data
- Fitness enthusiasts tracking workouts
- Users monitoring chronic conditions
- Anyone interested in wellness and preventive care

---

## Current Repository State

### Repository Structure
```
HealthTracker/
├── .git/                 # Git version control
├── LICENSE              # MIT License
└── CLAUDE.md           # This file
```

### Status
The repository is currently in its **initial setup phase** with minimal structure. Only the MIT License file exists.

### Next Steps for Development
1. **Choose Technology Stack** - Determine frontend/backend technologies
2. **Set Up Project Structure** - Create directory hierarchy
3. **Initialize Dependencies** - Set up package managers and dependencies
4. **Configure Development Environment** - Add linters, formatters, etc.
5. **Create Initial Documentation** - README.md, CONTRIBUTING.md

---

## Recommended Architecture

### Technology Stack Options

#### Option A: Full-Stack JavaScript/TypeScript
```
Frontend: React/Next.js + TypeScript
Backend: Node.js + Express/Fastify
Database: PostgreSQL/MongoDB
Mobile: React Native (optional)
```

#### Option B: Python Backend
```
Frontend: React/Vue.js + TypeScript
Backend: Python (FastAPI/Django)
Database: PostgreSQL
Mobile: Flutter/React Native (optional)
```

#### Option C: Modern Serverless
```
Frontend: Next.js/Nuxt.js
Backend: Serverless Functions (Vercel/AWS Lambda)
Database: Supabase/Firebase/PlanetScale
```

### Recommended Directory Structure

```
HealthTracker/
├── .github/              # GitHub Actions, templates
│   └── workflows/        # CI/CD pipelines
├── docs/                 # Additional documentation
├── src/
│   ├── api/             # Backend API code
│   │   ├── controllers/
│   │   ├── models/
│   │   ├── routes/
│   │   ├── middleware/
│   │   └── utils/
│   ├── client/          # Frontend application
│   │   ├── components/
│   │   ├── pages/
│   │   ├── hooks/
│   │   ├── services/
│   │   ├── styles/
│   │   └── utils/
│   └── shared/          # Shared types and utilities
├── tests/
│   ├── unit/
│   ├── integration/
│   └── e2e/
├── scripts/             # Build and deployment scripts
├── .env.example         # Environment variables template
├── .gitignore
├── package.json
├── tsconfig.json
├── README.md
├── CONTRIBUTING.md
├── CHANGELOG.md
├── LICENSE
└── CLAUDE.md           # This file
```

---

## Development Workflows

### Git Branching Strategy

#### Branch Naming Convention
```
feature/descriptive-name    # New features
bugfix/issue-description    # Bug fixes
hotfix/critical-fix         # Production hotfixes
docs/what-changed          # Documentation updates
refactor/what-refactored   # Code refactoring
test/what-tested           # Test additions
```

#### Workflow
1. Create feature branch from `main`
2. Make changes with clear, atomic commits
3. Write/update tests
4. Update documentation
5. Create Pull Request
6. Code review and approval
7. Merge to `main`

### Commit Message Convention

Follow **Conventional Commits** specification:

```
<type>(<scope>): <subject>

<body>

<footer>
```

**Types:**
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `style`: Code style changes (formatting)
- `refactor`: Code refactoring
- `test`: Adding/updating tests
- `chore`: Maintenance tasks
- `perf`: Performance improvements

**Examples:**
```
feat(api): add endpoint for blood pressure tracking
fix(ui): resolve chart rendering issue on mobile
docs(readme): update installation instructions
test(auth): add unit tests for login flow
```

---

## Code Conventions

### General Principles

1. **Write Clean, Readable Code**
   - Use meaningful variable and function names
   - Keep functions small and focused (single responsibility)
   - Add comments only when necessary to explain "why", not "what"

2. **Follow DRY (Don't Repeat Yourself)**
   - Extract reusable logic into functions/utilities
   - Create shared components for common UI elements

3. **Error Handling**
   - Always handle errors gracefully
   - Provide meaningful error messages
   - Log errors appropriately

4. **Type Safety**
   - Use TypeScript for type safety
   - Define interfaces for data structures
   - Avoid `any` type when possible

### Naming Conventions

#### JavaScript/TypeScript
```typescript
// Variables and functions: camelCase
const userName = "John";
function calculateBMI() {}

// Classes and Types: PascalCase
class HealthMetric {}
interface UserProfile {}
type MetricValue = number;

// Constants: UPPER_SNAKE_CASE
const MAX_HEART_RATE = 220;
const API_BASE_URL = "https://api.example.com";

// Files: kebab-case
// health-metrics.ts
// user-profile.service.ts
```

#### Python
```python
# Variables and functions: snake_case
user_name = "John"
def calculate_bmi():
    pass

# Classes: PascalCase
class HealthMetric:
    pass

# Constants: UPPER_SNAKE_CASE
MAX_HEART_RATE = 220
API_BASE_URL = "https://api.example.com"
```

### Code Organization

1. **Import Order**
   ```typescript
   // 1. External dependencies
   import React from 'react';
   import axios from 'axios';

   // 2. Internal modules
   import { HealthMetric } from '@/models';
   import { formatDate } from '@/utils';

   // 3. Relative imports
   import { Header } from '../components';

   // 4. Styles
   import styles from './styles.module.css';
   ```

2. **Component Structure (React)**
   ```typescript
   // 1. Imports
   // 2. Types/Interfaces
   // 3. Constants
   // 4. Component definition
   // 5. Styled components (if any)
   // 6. Export
   ```

---

## AI Assistant Guidelines

### When Working on This Project

#### 1. Always Check Context First
- Read relevant files before making changes
- Understand existing patterns and conventions
- Look for similar implementations in the codebase

#### 2. Plan Before Implementing
- Use TodoWrite tool for multi-step tasks
- Break down complex features into smaller tasks
- Communicate your plan before starting

#### 3. Follow Security Best Practices
- Never commit sensitive data (API keys, passwords)
- Validate and sanitize all user inputs
- Use parameterized queries for database operations
- Implement proper authentication and authorization
- Follow OWASP Top 10 security guidelines

#### 4. Write Tests
- Write unit tests for business logic
- Add integration tests for API endpoints
- Consider edge cases and error scenarios
- Aim for meaningful test coverage (not just high %)

#### 5. Documentation
- Update README.md when adding features
- Document API endpoints (use OpenAPI/Swagger)
- Add JSDoc/docstrings for complex functions
- Keep CLAUDE.md updated with architectural changes

#### 6. Code Quality
- Run linters and formatters before committing
- Fix all TypeScript/linting errors
- Ensure code builds successfully
- Test changes locally before pushing

#### 7. Git Practices
- Write clear, descriptive commit messages
- Keep commits atomic and focused
- Don't commit commented-out code
- Use `.gitignore` properly

#### 8. Communication
- Ask clarifying questions when requirements are unclear
- Explain complex changes in PR descriptions
- Highlight breaking changes
- Suggest alternatives when appropriate

### Common Tasks

#### Adding a New Health Metric
1. Define the data model/schema
2. Create database migration (if applicable)
3. Add API endpoints (CRUD operations)
4. Implement frontend components
5. Add validation logic
6. Write tests
7. Update documentation

#### Creating a New API Endpoint
1. Define route in appropriate router file
2. Create controller function
3. Add validation middleware
4. Implement business logic
5. Add error handling
6. Write tests
7. Document in API documentation

#### Adding a UI Component
1. Create component file in appropriate directory
2. Implement component logic
3. Add styling
4. Write prop types/interfaces
5. Add to component library/exports
6. Write tests
7. Add to Storybook (if used)

---

## Security Considerations

### Critical Security Rules

1. **Never Store Sensitive Data in Code**
   - Use environment variables for secrets
   - Add `.env` to `.gitignore`
   - Provide `.env.example` with dummy values

2. **Input Validation**
   - Validate all user inputs on both client and server
   - Use schema validation libraries (Joi, Zod, etc.)
   - Sanitize data before storage

3. **Authentication & Authorization**
   - Use established libraries (Passport.js, Auth0, etc.)
   - Implement JWT or session-based auth properly
   - Validate tokens on protected routes
   - Use HTTPS in production

4. **Database Security**
   - Use parameterized queries/ORMs
   - Never concatenate SQL strings
   - Implement proper access controls
   - Encrypt sensitive data at rest

5. **Health Data Privacy**
   - Comply with HIPAA (if applicable in US)
   - Follow GDPR guidelines (if serving EU users)
   - Implement data encryption
   - Provide data export/deletion capabilities
   - Add privacy policy and terms of service

6. **API Security**
   - Implement rate limiting
   - Use CORS properly
   - Validate request origins
   - Log security events

---

## Testing Strategy

### Test Pyramid

```
        /\
       /E2E\         <- Few, high-value end-to-end tests
      /------\
     /  Intg  \      <- Integration tests for APIs, components
    /----------\
   /    Unit    \    <- Many unit tests for logic
  /--------------\
```

### Unit Tests
- Test individual functions and methods
- Mock external dependencies
- Fast execution
- High coverage for business logic

### Integration Tests
- Test API endpoints
- Test database operations
- Test component integration
- Use test database

### End-to-End Tests
- Test critical user flows
- Test in browser environment
- Use tools like Cypress, Playwright
- Run in CI/CD pipeline

### Test File Naming
```
src/utils/date-formatter.ts
tests/unit/utils/date-formatter.test.ts

src/api/controllers/metrics.controller.ts
tests/integration/api/metrics.api.test.ts
```

---

## Development Environment Setup

### Prerequisites (To Be Determined)
Once the technology stack is chosen, document:
- Required software versions (Node.js, Python, etc.)
- Database setup instructions
- Environment variables needed
- IDE recommendations and extensions

### Getting Started (Template)
```bash
# Clone the repository
git clone <repo-url>
cd HealthTracker

# Install dependencies
# (commands will vary based on tech stack)

# Set up environment variables
cp .env.example .env
# Edit .env with your values

# Run database migrations
# (if applicable)

# Start development server
# (command TBD)
```

---

## API Design Guidelines (Future)

### RESTful Conventions
```
GET    /api/metrics              # List all metrics
GET    /api/metrics/:id          # Get specific metric
POST   /api/metrics              # Create new metric
PUT    /api/metrics/:id          # Update metric
DELETE /api/metrics/:id          # Delete metric
```

### Response Format
```json
{
  "success": true,
  "data": {},
  "message": "Operation successful",
  "timestamp": "2025-11-16T12:00:00Z"
}
```

### Error Response Format
```json
{
  "success": false,
  "error": {
    "code": "VALIDATION_ERROR",
    "message": "Invalid input data",
    "details": []
  },
  "timestamp": "2025-11-16T12:00:00Z"
}
```

---

## Data Models (Future)

### Core Entities (Suggestions)

1. **User**
   - id, email, password_hash, name, date_of_birth, gender
   - created_at, updated_at

2. **HealthMetric**
   - id, user_id, metric_type, value, unit, timestamp
   - notes, created_at

3. **Activity**
   - id, user_id, activity_type, duration, distance
   - calories_burned, timestamp, notes

4. **Goal**
   - id, user_id, goal_type, target_value, deadline
   - status, created_at, achieved_at

5. **Measurement**
   - id, user_id, measurement_type, value, unit, timestamp

---

## Performance Considerations

### Database
- Index frequently queried fields
- Use connection pooling
- Implement caching where appropriate
- Optimize queries (avoid N+1 problems)

### Frontend
- Code splitting and lazy loading
- Optimize images and assets
- Minimize bundle size
- Use memoization for expensive computations

### API
- Implement pagination for list endpoints
- Use compression for responses
- Cache static resources
- Rate limiting to prevent abuse

---

## Accessibility

Ensure the application is accessible:
- Use semantic HTML
- Provide alt text for images
- Ensure keyboard navigation works
- Maintain sufficient color contrast
- Support screen readers
- Follow WCAG 2.1 guidelines

---

## Deployment (Future)

Document deployment process once determined:
- Hosting platform
- CI/CD pipeline setup
- Environment configuration
- Database hosting
- Monitoring and logging
- Backup strategy

---

## Changelog

### 2025-11-16 - Initial Creation
- Created CLAUDE.md with comprehensive guidelines
- Documented repository structure and conventions
- Added AI assistant guidelines
- Established security and testing strategies

---

## Questions or Issues?

When you need guidance:
1. Check this CLAUDE.md file first
2. Look for similar patterns in the codebase
3. Ask the project maintainer for clarification
4. Document decisions for future reference

---

## Useful Resources

- [Conventional Commits](https://www.conventionalcommits.org/)
- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
- [HIPAA Compliance Guide](https://www.hhs.gov/hipaa/index.html)
- [GDPR Overview](https://gdpr.eu/)
- [REST API Best Practices](https://restfulapi.net/)

---

**Remember:** This document is a living guide. Update it as the project evolves, patterns emerge, and decisions are made. Keep it accurate and relevant to help all contributors (human and AI) work effectively on HealthTracker.
