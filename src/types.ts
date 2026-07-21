export interface NovoUser {
  user_id: string;
  email: string;
  org: string | null;
  tier: string;
  credits_available?: number;
}

export interface MoleculeProfile {
  smiles: string;
  source: 'enriched_database' | 'computed+admet';
  in_database: boolean;
  properties: Record<string, number | string | null>;
  admet?: Record<string, number | string | null> | null;
  compliance: Record<string, unknown>;
  structural_alerts?: Record<string, unknown>;
}

export interface ApiUsage {
  credits?: number;
  credits_remaining?: number;
  credit_status?: 'ok' | 'low' | 'exhausted';
  credit_warning?: { credits_remaining: number; message: string; upgrade_url: string };
  tool?: string;
  source?: string;
  funnel_id?: string;
}

export interface ApiResponse<T> {
  result: T;
  usage: ApiUsage;
}

export interface ApiError {
  error: string;
  error_code?: string;
  message?: string;
  upgrade_url?: string;
  [key: string]: unknown;
}

// Compute-tier result shapes — loose, since the renderer probes multiple field aliases.
export interface PkaResult { pka_values?: number[]; ionizable_groups?: string[]; interpretation?: string; [k: string]: unknown; }
export interface SolubilityResult { logS?: number; solubility_mg_ml?: number; category?: string; [k: string]: unknown; }
export interface BdeResult { bonds?: Array<{ atoms: string; bde_kcal_mol: number }>; weakest_bond?: { atoms: string; bde_kcal_mol: number }; interpretation?: string; bond_count?: number; [k: string]: unknown; }
export interface FrontierOrbitalsResult { homo_ev?: number; lumo_ev?: number; gap_ev?: number; [k: string]: unknown; }
