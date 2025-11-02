import { SignIn } from '@stackframe/stack';

export default function SignInPage() {
  return (
    <div className="flex items-center justify-center min-h-screen bg-white">
      <div className="w-full max-w-md p-8">
        <div className="mb-8 text-center">
          <div className="text-5xl mb-4">📊</div>
          <h1 className="text-3xl font-bold text-gray-900 mb-2">SheetGrid</h1>
          <p className="text-gray-600">Sign in to continue</p>
        </div>
        <SignIn />
      </div>
    </div>
  );
}

