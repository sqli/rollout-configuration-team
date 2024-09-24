"use client";

import { useSession, signIn, signOut } from 'next-auth/react';
import Dashboard from '../components/Dashboard';
import { NextUIProvider } from '@nextui-org/react';

export default function Home() {
  const { data: session, status } = useSession();
  const loading = status === 'loading';

  if (loading) return <p>Loading...</p>;

  return (
    <NextUIProvider>
      <div>
      {!true ? (
        <>
          <h1>You are not signed in</h1>
          <button
            onClick={() =>
              signIn('azure-ad')
            }
          >
            Sign in with Microsoft
          </button>
        </>
      ) : (
        <>
          {/* <h1>Welcome, {session.user?.name}</h1> */}
          <Dashboard />

          {/* <button onClick={() => signOut()}>Sign out</button> */}
        </>
      )}
    </div>
    </NextUIProvider>
    
  );
}
