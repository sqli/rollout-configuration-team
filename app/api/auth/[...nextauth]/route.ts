import NextAuth from 'next-auth';
import AzureADProvider from 'next-auth/providers/azure-ad';

const env = process.env;

const handler = NextAuth({
  providers: [
    AzureADProvider({
      clientId: env.NEXT_PUBLIC_AZURE_AD_CLIENT_ID || "",
      clientSecret: env.NEXT_PUBLIC_AZURE_AD_CLIENT_SECRET || '',
      tenantId: env.NEXT_PUBLIC_AZURE_AD_TENANT_ID,
      authorization: {
        params: { 
            scope: 'openid profile email offline_access',
           
        },
      },
    }),
  ],
  callbacks: {
    // async signIn({ profile }) {
    //   if (profile?.email?.endsWith('@yourcompany.com')) {
    //     return true;
    //   } else {
    //     return false;
    //   }
    // },
    async jwt({ token, account }) {
      if (account) {
        token.accessToken = account.access_token;
      }
      return token;
    },
    async session({ session, token }:any) {
      session.accessToken = token.accessToken;
      return session;
    },
  },
});

export { handler as GET, handler as POST };
