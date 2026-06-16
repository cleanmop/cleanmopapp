import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://ipojirqxrulicgmkjbpm.supabase.co";
const supabaseKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imlwb2ppcnF4cnVsaWNnbWtqYnBtIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODE1ODAxNTUsImV4cCI6MjA5NzE1NjE1NX0.A5p44m9x0rQebOeWcMqWTCV1rs34Nm3JQh5G2Ydw3hE";

export const supabase = createClient(supabaseUrl, supabaseKey);