import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL || 'https://qexahlsjnbugzbdvhysm.supabase.co';
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || 'sb_publishable_pQn_-Je4yCTm4UVznd94Iw_FBbFug8o';

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
