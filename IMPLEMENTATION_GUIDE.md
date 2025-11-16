# HealthTracker Implementation Guide

**Quick Start Guide for Developers**

> This is a condensed guide. See [ARCHITECTURE_REVIEW.md](./ARCHITECTURE_REVIEW.md) for complete details.

---

## 🎯 Architecture Summary

**Pattern:** Serverless + JAMstack
**Stack:** Next.js + TypeScript + PostgreSQL + Redis
**Deployment:** Vercel + Supabase
**Mobile:** PWA (Progressive Web App)

---

## 🚀 Quick Start (Getting Started)

### 1. Initial Setup

```bash
# Create Next.js project
npx create-next-app@latest healthtracker \
  --typescript \
  --tailwind \
  --app \
  --use-npm

cd healthtracker

# Install core dependencies
npm install @supabase/supabase-js \
  next-auth \
  zod \
  @tanstack/react-query \
  zustand \
  recharts \
  date-fns

# Install dev dependencies
npm install -D @types/node \
  eslint \
  prettier \
  husky \
  lint-staged \
  vitest \
  @testing-library/react \
  @playwright/test
```

### 2. Project Structure

```bash
mkdir -p src/{app,components,lib,types,hooks,services}
mkdir -p src/app/{api,auth,dashboard,metrics,activities,nutrition,goals,analytics}
mkdir -p tests/{unit,integration,e2e}
mkdir -p scripts
```

### 3. Environment Setup

```bash
# Create .env.local
cat > .env.local << 'EOF'
# Database
DATABASE_URL="postgresql://..."
DIRECT_URL="postgresql://..."

# Supabase
NEXT_PUBLIC_SUPABASE_URL="https://xxx.supabase.co"
NEXT_PUBLIC_SUPABASE_ANON_KEY="..."
SUPABASE_SERVICE_ROLE_KEY="..."

# NextAuth
NEXTAUTH_URL="http://localhost:3000"
NEXTAUTH_SECRET="..." # Generate with: openssl rand -base64 32

# OAuth (optional)
GOOGLE_CLIENT_ID="..."
GOOGLE_CLIENT_SECRET="..."

# Redis
UPSTASH_REDIS_REST_URL="..."
UPSTASH_REDIS_REST_TOKEN="..."

# Monitoring
SENTRY_DSN="..."
NEXT_PUBLIC_SENTRY_DSN="..."

# Environment
NODE_ENV="development"
EOF

# Add to .gitignore
echo ".env.local" >> .gitignore
```

### 4. Database Setup

```bash
# Install Prisma
npm install prisma @prisma/client
npx prisma init

# Create schema (see ARCHITECTURE_REVIEW.md for full schema)
# Then run migrations
npx prisma migrate dev --name init
npx prisma generate
```

---

## 📁 Recommended Directory Structure

```
healthtracker/
├── .github/
│   └── workflows/
│       ├── ci.yml
│       └── deploy.yml
├── public/
│   ├── icons/
│   ├── manifest.json
│   └── sw.js
├── src/
│   ├── app/                    # Next.js 14 App Router
│   │   ├── (auth)/
│   │   │   ├── login/
│   │   │   ├── register/
│   │   │   └── layout.tsx
│   │   ├── (dashboard)/
│   │   │   ├── dashboard/
│   │   │   ├── metrics/
│   │   │   ├── activities/
│   │   │   ├── nutrition/
│   │   │   ├── goals/
│   │   │   ├── analytics/
│   │   │   └── layout.tsx
│   │   ├── api/
│   │   │   ├── auth/
│   │   │   ├── metrics/
│   │   │   ├── activities/
│   │   │   ├── nutrition/
│   │   │   ├── goals/
│   │   │   └── analytics/
│   │   ├── layout.tsx
│   │   └── page.tsx
│   ├── components/
│   │   ├── ui/                # shadcn/ui components
│   │   ├── forms/
│   │   ├── charts/
│   │   ├── layouts/
│   │   └── shared/
│   ├── lib/
│   │   ├── db.ts              # Database client
│   │   ├── auth.ts            # Auth configuration
│   │   ├── redis.ts           # Redis client
│   │   ├── validations/       # Zod schemas
│   │   └── utils.ts
│   ├── hooks/
│   │   ├── useMetrics.ts
│   │   ├── useActivities.ts
│   │   └── useAuth.ts
│   ├── services/
│   │   ├── metrics.service.ts
│   │   ├── activities.service.ts
│   │   └── analytics.service.ts
│   └── types/
│       ├── models.ts
│       ├── api.ts
│       └── index.ts
├── tests/
│   ├── unit/
│   ├── integration/
│   └── e2e/
├── scripts/
│   ├── seed.ts
│   └── migrate.ts
├── .env.example
├── .env.local
├── .eslintrc.json
├── .prettierrc
├── next.config.js
├── tsconfig.json
├── package.json
├── prisma/
│   └── schema.prisma
├── ARCHITECTURE_REVIEW.md
├── CLAUDE.md
├── README.md
└── LICENSE
```

---

## 🔧 Configuration Files

### next.config.js

```javascript
/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  swcMinify: true,
  images: {
    domains: ['supabase.co'],
    formats: ['image/avif', 'image/webp'],
  },
  experimental: {
    serverActions: true,
  },
};

module.exports = nextConfig;
```

### tsconfig.json

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "lib": ["dom", "dom.iterable", "esnext"],
    "allowJs": true,
    "skipLibCheck": true,
    "strict": true,
    "forceConsistentCasingInFileNames": true,
    "noEmit": true,
    "esModuleInterop": true,
    "module": "esnext",
    "moduleResolution": "bundler",
    "resolveJsonModule": true,
    "isolatedModules": true,
    "jsx": "preserve",
    "incremental": true,
    "plugins": [
      {
        "name": "next"
      }
    ],
    "paths": {
      "@/*": ["./src/*"]
    }
  },
  "include": ["next-env.d.ts", "**/*.ts", "**/*.tsx", ".next/types/**/*.ts"],
  "exclude": ["node_modules"]
}
```

### .eslintrc.json

```json
{
  "extends": [
    "next/core-web-vitals",
    "plugin:@typescript-eslint/recommended",
    "prettier"
  ],
  "rules": {
    "@typescript-eslint/no-unused-vars": "error",
    "@typescript-eslint/no-explicit-any": "warn",
    "no-console": ["warn", { "allow": ["warn", "error"] }]
  }
}
```

### .prettierrc

```json
{
  "semi": true,
  "trailingComma": "es5",
  "singleQuote": true,
  "printWidth": 80,
  "tabWidth": 2,
  "useTabs": false
}
```

---

## 🗄️ Database Schema (Prisma)

```prisma
// prisma/schema.prisma
generator client {
  provider = "prisma-client-js"
}

datasource db {
  provider = "postgresql"
  url      = env("DATABASE_URL")
  directUrl = env("DIRECT_URL")
}

model User {
  id            String    @id @default(uuid())
  email         String    @unique
  emailVerified Boolean   @default(false)
  passwordHash  String
  firstName     String?
  lastName      String?
  dateOfBirth   DateTime?
  gender        String?
  height        Decimal?
  units         String    @default("metric")
  timezone      String?
  language      String    @default("en")
  avatarUrl     String?
  role          String    @default("user")
  isActive      Boolean   @default(true)
  createdAt     DateTime  @default(now())
  updatedAt     DateTime  @updatedAt
  lastLoginAt   DateTime?

  metrics    HealthMetric[]
  activities Activity[]
  nutrition  NutritionEntry[]
  goals      Goal[]

  @@map("users")
}

model HealthMetric {
  id         String   @id @default(uuid())
  userId     String
  metricType String
  value      Decimal
  unit       String
  timestamp  DateTime
  source     String   @default("manual")
  deviceId   String?
  notes      String?
  tags       String[]
  createdAt  DateTime @default(now())

  user User @relation(fields: [userId], references: [id], onDelete: Cascade)

  @@index([userId])
  @@index([timestamp])
  @@index([metricType])
  @@index([userId, metricType, timestamp])
  @@map("health_metrics")
}

model Activity {
  id               String   @id @default(uuid())
  userId           String
  activityType     String
  name             String
  startTime        DateTime
  endTime          DateTime
  duration         Int
  distance         Decimal?
  caloriesBurned   Int?
  averageHeartRate Int?
  maxHeartRate     Int?
  intensity        String?
  notes            String?
  route            Json?
  createdAt        DateTime @default(now())

  user User @relation(fields: [userId], references: [id], onDelete: Cascade)

  @@index([userId])
  @@index([startTime])
  @@map("activities")
}

model NutritionEntry {
  id            String   @id @default(uuid())
  userId        String
  mealType      String
  timestamp     DateTime
  foods         Json
  totalCalories Int?
  totalProtein  Decimal?
  totalCarbs    Decimal?
  totalFat      Decimal?
  notes         String?
  createdAt     DateTime @default(now())

  user User @relation(fields: [userId], references: [id], onDelete: Cascade)

  @@index([userId])
  @@index([timestamp])
  @@map("nutrition_entries")
}

model Goal {
  id           String    @id @default(uuid())
  userId       String
  goalType     String
  title        String
  description  String?
  targetValue  Decimal
  currentValue Decimal   @default(0)
  unit         String?
  startDate    DateTime
  targetDate   DateTime
  status       String    @default("active")
  progress     Int       @default(0)
  milestones   Json?
  createdAt    DateTime  @default(now())
  updatedAt    DateTime  @updatedAt
  completedAt  DateTime?

  user User @relation(fields: [userId], references: [id], onDelete: Cascade)

  @@index([userId])
  @@index([status])
  @@map("goals")
}
```

---

## 🔐 Authentication Setup

```typescript
// src/lib/auth.ts
import { NextAuthOptions } from 'next-auth';
import CredentialsProvider from 'next-auth/providers/credentials';
import GoogleProvider from 'next-auth/providers/google';
import { PrismaAdapter } from '@next-auth/prisma-adapter';
import { compare } from 'bcryptjs';
import { prisma } from './db';

export const authOptions: NextAuthOptions = {
  adapter: PrismaAdapter(prisma),
  session: {
    strategy: 'jwt',
    maxAge: 7 * 24 * 60 * 60, // 7 days
  },
  providers: [
    CredentialsProvider({
      name: 'credentials',
      credentials: {
        email: { label: 'Email', type: 'email' },
        password: { label: 'Password', type: 'password' },
      },
      async authorize(credentials) {
        if (!credentials?.email || !credentials?.password) {
          return null;
        }

        const user = await prisma.user.findUnique({
          where: { email: credentials.email },
        });

        if (!user || !user.passwordHash) {
          return null;
        }

        const isPasswordValid = await compare(
          credentials.password,
          user.passwordHash
        );

        if (!isPasswordValid) {
          return null;
        }

        return {
          id: user.id,
          email: user.email,
          name: `${user.firstName} ${user.lastName}`,
        };
      },
    }),
    GoogleProvider({
      clientId: process.env.GOOGLE_CLIENT_ID!,
      clientSecret: process.env.GOOGLE_CLIENT_SECRET!,
    }),
  ],
  callbacks: {
    async jwt({ token, user }) {
      if (user) {
        token.id = user.id;
      }
      return token;
    },
    async session({ session, token }) {
      if (session.user) {
        session.user.id = token.id as string;
      }
      return session;
    },
  },
  pages: {
    signIn: '/login',
    signOut: '/logout',
    error: '/error',
  },
};
```

---

## 📊 Example API Route

```typescript
// src/app/api/metrics/route.ts
import { NextRequest, NextResponse } from 'next/server';
import { getServerSession } from 'next-auth';
import { z } from 'zod';
import { prisma } from '@/lib/db';
import { authOptions } from '@/lib/auth';

const createMetricSchema = z.object({
  metricType: z.enum(['weight', 'blood_pressure', 'heart_rate']),
  value: z.number().positive(),
  unit: z.string(),
  timestamp: z.string().datetime(),
  notes: z.string().optional(),
});

export async function GET(req: NextRequest) {
  const session = await getServerSession(authOptions);

  if (!session?.user?.id) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  const { searchParams } = new URL(req.url);
  const metricType = searchParams.get('type');
  const limit = parseInt(searchParams.get('limit') || '50');

  const metrics = await prisma.healthMetric.findMany({
    where: {
      userId: session.user.id,
      ...(metricType && { metricType }),
    },
    orderBy: { timestamp: 'desc' },
    take: limit,
  });

  return NextResponse.json({ success: true, data: metrics });
}

export async function POST(req: NextRequest) {
  const session = await getServerSession(authOptions);

  if (!session?.user?.id) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  try {
    const body = await req.json();
    const validatedData = createMetricSchema.parse(body);

    const metric = await prisma.healthMetric.create({
      data: {
        userId: session.user.id,
        metricType: validatedData.metricType,
        value: validatedData.value,
        unit: validatedData.unit,
        timestamp: new Date(validatedData.timestamp),
        notes: validatedData.notes,
      },
    });

    return NextResponse.json({ success: true, data: metric }, { status: 201 });
  } catch (error) {
    if (error instanceof z.ZodError) {
      return NextResponse.json(
        { success: false, error: error.errors },
        { status: 400 }
      );
    }
    return NextResponse.json(
      { success: false, error: 'Internal server error' },
      { status: 500 }
    );
  }
}
```

---

## 🎨 Example Component

```typescript
// src/components/metrics/MetricChart.tsx
'use client';

import { useMemo } from 'react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from 'recharts';
import { format } from 'date-fns';

interface MetricChartProps {
  data: Array<{
    timestamp: Date;
    value: number;
  }>;
  metricType: string;
  unit: string;
}

export function MetricChart({ data, metricType, unit }: MetricChartProps) {
  const chartData = useMemo(() => {
    return data.map((item) => ({
      date: format(new Date(item.timestamp), 'MMM dd'),
      value: item.value,
    }));
  }, [data]);

  return (
    <div className="w-full h-80">
      <ResponsiveContainer width="100%" height="100%">
        <LineChart data={chartData}>
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis dataKey="date" />
          <YAxis label={{ value: unit, angle: -90, position: 'insideLeft' }} />
          <Tooltip />
          <Line
            type="monotone"
            dataKey="value"
            stroke="#3b82f6"
            strokeWidth={2}
            dot={{ r: 4 }}
          />
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
}
```

---

## 🧪 Testing Setup

```typescript
// vitest.config.ts
import { defineConfig } from 'vitest/config';
import react from '@vitejs/plugin-react';
import path from 'path';

export default defineConfig({
  plugins: [react()],
  test: {
    environment: 'jsdom',
    globals: true,
    setupFiles: ['./tests/setup.ts'],
  },
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
    },
  },
});
```

```typescript
// tests/unit/lib/utils.test.ts
import { describe, it, expect } from 'vitest';
import { calculateBMI } from '@/lib/utils';

describe('calculateBMI', () => {
  it('should calculate BMI correctly', () => {
    expect(calculateBMI(70, 1.75)).toBeCloseTo(22.86, 1);
  });

  it('should throw error for invalid weight', () => {
    expect(() => calculateBMI(-70, 1.75)).toThrow();
  });
});
```

---

## 📝 Key Implementation Steps

### Week 1-2: Foundation
1. ✅ Set up Next.js project with TypeScript
2. ✅ Configure Supabase database
3. ✅ Set up authentication with NextAuth.js
4. ✅ Create basic UI components with shadcn/ui
5. ✅ Set up CI/CD with GitHub Actions

### Week 3-4: Core Features
1. ✅ Implement health metrics CRUD
2. ✅ Build data visualization components
3. ✅ Create user dashboard
4. ✅ Add activity tracking
5. ✅ Implement responsive design

### Week 5-6: Advanced Features
1. ✅ Add nutrition tracking
2. ✅ Implement goal setting
3. ✅ Build analytics engine
4. ✅ Add data export functionality
5. ✅ Implement PWA features

### Week 7-8: Polish
1. ✅ Performance optimization
2. ✅ Security audit
3. ✅ Comprehensive testing
4. ✅ Documentation
5. ✅ Beta launch preparation

---

## 🚦 Development Commands

```bash
# Development
npm run dev              # Start dev server
npm run build           # Build for production
npm run start           # Start production server
npm run lint            # Run ESLint
npm run format          # Run Prettier
npm run type-check      # TypeScript check

# Database
npx prisma studio       # Open Prisma Studio
npx prisma migrate dev  # Run migrations
npx prisma generate     # Generate Prisma Client

# Testing
npm run test            # Run unit tests
npm run test:e2e        # Run E2E tests
npm run test:coverage   # Generate coverage report

# Deployment
vercel                  # Deploy to Vercel
```

---

## 🔗 Useful Links

- **Architecture Details:** [ARCHITECTURE_REVIEW.md](./ARCHITECTURE_REVIEW.md)
- **AI Guidelines:** [CLAUDE.md](./CLAUDE.md)
- **Next.js Docs:** https://nextjs.org/docs
- **Prisma Docs:** https://www.prisma.io/docs
- **Supabase Docs:** https://supabase.com/docs
- **shadcn/ui:** https://ui.shadcn.com

---

## ❓ Need Help?

1. Check [CLAUDE.md](./CLAUDE.md) for development guidelines
2. Review [ARCHITECTURE_REVIEW.md](./ARCHITECTURE_REVIEW.md) for architecture decisions
3. Consult the documentation links above
4. Create an issue in the repository

---

**Happy Coding!** 🎉
