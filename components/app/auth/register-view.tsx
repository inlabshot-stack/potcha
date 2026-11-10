'use client';

import { useState } from 'react';
import { toast } from 'sonner';
import { ArrowLeft, Building2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import {
  Select, SelectContent, SelectItem, SelectTrigger, SelectValue,
} from '@/components/ui/select';
import { api, ApiClientError, saveSessionToken } from '@/lib/client/api';
import type { SessionUser } from '@/lib/client/types';
import { useAppStore } from '@/components/app/store';
import { navigate } from '@/components/app/router';
import { CURRENCIES } from '@/lib/constants';

export function RegisterView() {
  const setUser = useAppStore((s) => s.setUser);
  const [form, setForm] = useState({
    companyName: '', name: '', email: '', password: '', currency: 'EGP',
  });
  const [busy, setBusy] = useState(false);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true);
    try {
      const res = await api<{ sessionToken?: string }>('/auth/register', { method: 'POST', body: form });
      if (res.sessionToken) saveSessionToken(res.sessionToken);
      const me = await api<{ user: SessionUser }>('/auth/me');
      setUser(me.user);
      toast.success('Company workspace created', {
        description: `Welcome to Potcha Ops, ${me.user.name}. You are the CEO of ${me.user.companyName}.`,
      });
      navigate('/dashboard');
    } catch (err) {
      toast.error(err instanceof ApiClientError ? err.message : 'Registration failed');
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="flex min-h-screen flex-col bg-background">
      <div className="flex flex-1 items-center justify-center px-4 py-10">
        <div className="w-full max-w-lg">
          <Button variant="ghost" size="sm" className="mb-4 gap-1.5" onClick={() => navigate('/dashboard')}>
            <ArrowLeft className="h-4 w-4" aria-hidden /> Back to sign in
          </Button>
          <Card>
            <CardHeader>
              <div className="mb-1 flex h-10 w-10 items-center justify-center rounded-xl bg-primary/10">
                <Building2 className="h-5 w-5 text-primary" aria-hidden />
              </div>
              <CardTitle className="text-lg">Create your company workspace</CardTitle>
              <CardDescription>
                You will be the CEO (Level 3) with full oversight. Your data stays isolated
                from every other company — nothing is shared.
              </CardDescription>
            </CardHeader>
            <CardContent>
              <form className="space-y-4" onSubmit={submit}>
                <div className="grid gap-4 sm:grid-cols-2">
                  <div className="space-y-1.5 sm:col-span-2">
                    <Label htmlFor="reg-company">Company name</Label>
                    <Input
                      id="reg-company" placeholder="e.g. Nile Rose Weddings"
                      value={form.companyName}
                      onChange={(e) => setForm((f) => ({ ...f, companyName: e.target.value }))}
                      required minLength={2}
                    />
                  </div>
                  <div className="space-y-1.5">
                    <Label htmlFor="reg-name">Your name</Label>
                    <Input
                      id="reg-name" placeholder="e.g. Amira Hassan"
                      value={form.name}
                      onChange={(e) => setForm((f) => ({ ...f, name: e.target.value }))}
                      required minLength={2}
                    />
                  </div>
                  <div className="space-y-1.5">
                    <Label htmlFor="reg-currency">Currency</Label>
                    <Select
                      value={form.currency}
                      onValueChange={(v) => setForm((f) => ({ ...f, currency: v }))}
                    >
                      <SelectTrigger id="reg-currency"><SelectValue /></SelectTrigger>
                      <SelectContent>
                        {CURRENCIES.map((c) => (
                          <SelectItem key={c} value={c}>{c}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-1.5 sm:col-span-2">
                    <Label htmlFor="reg-email">Work email</Label>
                    <Input
                      id="reg-email" type="email" placeholder="ceo@yourcompany.com"
                      autoComplete="email"
                      value={form.email}
                      onChange={(e) => setForm((f) => ({ ...f, email: e.target.value }))}
                      required
                    />
                  </div>
                  <div className="space-y-1.5 sm:col-span-2">
                    <Label htmlFor="reg-password">Password</Label>
                    <Input
                      id="reg-password" type="password" placeholder="At least 8 chars, letters + numbers"
                      autoComplete="new-password"
                      value={form.password}
                      onChange={(e) => setForm((f) => ({ ...f, password: e.target.value }))}
                      required minLength={8}
                    />
                    <p className="text-[11px] text-muted-foreground">
                      Stored as a bcrypt hash — never in plain text.
                    </p>
                  </div>
                </div>
                <Button type="submit" className="w-full" disabled={busy}>
                  {busy ? 'Creating workspace…' : 'Create workspace'}
                </Button>
              </form>
            </CardContent>
          </Card>
        </div>
      </div>
      <footer className="mt-auto border-t bg-background px-4 py-3 pb-[calc(0.75rem+env(safe-area-inset-bottom))] text-center text-[11px] text-muted-foreground">
        Potcha Ops · Every Wedding. Every Detail. Under Control.
      </footer>
    </div>
  );
}
