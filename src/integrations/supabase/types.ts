export type Json =
  | string
  | number
  | boolean
  | null
  | { [key: string]: Json | undefined }
  | Json[]

export type Database = {
  // Allows to automatically instantiate createClient with right options
  // instead of createClient<Database, { PostgrestVersion: 'XX' }>(URL, KEY)
  __InternalSupabase: {
    PostgrestVersion: "14.4"
  }
  public: {
    Tables: {
      financial_categories: {
        Row: {
          code: string
          created_at: string
          full_label: string
          id: string
          name: string
          tokens: string[]
          updated_at: string
        }
        Insert: {
          code: string
          created_at?: string
          full_label: string
          id?: string
          name: string
          tokens?: string[]
          updated_at?: string
        }
        Update: {
          code?: string
          created_at?: string
          full_label?: string
          id?: string
          name?: string
          tokens?: string[]
          updated_at?: string
        }
        Relationships: []
      }
      import_templates: {
        Row: {
          created_at: string
          id: string
          import_type: string
          is_default: boolean | null
          name: string
          template_storage_path: string
          version: number | null
        }
        Insert: {
          created_at?: string
          id?: string
          import_type: string
          is_default?: boolean | null
          name: string
          template_storage_path: string
          version?: number | null
        }
        Update: {
          created_at?: string
          id?: string
          import_type?: string
          is_default?: boolean | null
          name?: string
          template_storage_path?: string
          version?: number | null
        }
        Relationships: []
      }
      runs: {
        Row: {
          created_at: string
          error_report_json: Json | null
          id: string
          import_type: Database["public"]["Enums"]["import_type"]
          output_csv_path: string | null
          output_storage_path: string | null
          output_xlsx_path: string | null
          preview_json: Json | null
          status: Database["public"]["Enums"]["run_status"]
          transform_instructions: string | null
          upload_id: string
        }
        Insert: {
          created_at?: string
          error_report_json?: Json | null
          id?: string
          import_type: Database["public"]["Enums"]["import_type"]
          output_csv_path?: string | null
          output_storage_path?: string | null
          output_xlsx_path?: string | null
          preview_json?: Json | null
          status?: Database["public"]["Enums"]["run_status"]
          transform_instructions?: string | null
          upload_id: string
        }
        Update: {
          created_at?: string
          error_report_json?: Json | null
          id?: string
          import_type?: Database["public"]["Enums"]["import_type"]
          output_csv_path?: string | null
          output_storage_path?: string | null
          output_xlsx_path?: string | null
          preview_json?: Json | null
          status?: Database["public"]["Enums"]["run_status"]
          transform_instructions?: string | null
          upload_id?: string
        }
        Relationships: [
          {
            foreignKeyName: "runs_upload_id_fkey"
            columns: ["upload_id"]
            isOneToOne: false
            referencedRelation: "uploads"
            referencedColumns: ["id"]
          },
        ]
      }
      uploads: {
        Row: {
          created_at: string
          error_message: string | null
          id: string
          original_filename: string
          status: Database["public"]["Enums"]["upload_status"]
          storage_path: string
        }
        Insert: {
          created_at?: string
          error_message?: string | null
          id?: string
          original_filename: string
          status?: Database["public"]["Enums"]["upload_status"]
          storage_path: string
        }
        Update: {
          created_at?: string
          error_message?: string | null
          id?: string
          original_filename?: string
          status?: Database["public"]["Enums"]["upload_status"]
          storage_path?: string
        }
        Relationships: []
      }
    }
    Views: {
      [_ in never]: never
    }
    Functions: {
      [_ in never]: never
    }
    Enums: {
      import_type:
        | "clientes_fornecedores"
        | "movimentacoes"
        | "vendas"
        | "contratos"
      run_status: "queued" | "processing" | "done" | "error"
      upload_status: "uploaded" | "processing" | "done" | "error"
    }
    CompositeTypes: {
      [_ in never]: never
    }
  }
}

type DatabaseWithoutInternals = Omit<Database, "__InternalSupabase">

type DefaultSchema = DatabaseWithoutInternals[Extract<keyof Database, "public">]

export type Tables<
  DefaultSchemaTableNameOrOptions extends
    | keyof (DefaultSchema["Tables"] & DefaultSchema["Views"])
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
        DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
      DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])[TableName] extends {
      Row: infer R
    }
    ? R
    : never
  : DefaultSchemaTableNameOrOptions extends keyof (DefaultSchema["Tables"] &
        DefaultSchema["Views"])
    ? (DefaultSchema["Tables"] &
        DefaultSchema["Views"])[DefaultSchemaTableNameOrOptions] extends {
        Row: infer R
      }
      ? R
      : never
    : never

export type TablesInsert<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Insert: infer I
    }
    ? I
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Insert: infer I
      }
      ? I
      : never
    : never

export type TablesUpdate<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Update: infer U
    }
    ? U
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Update: infer U
      }
      ? U
      : never
    : never

export type Enums<
  DefaultSchemaEnumNameOrOptions extends
    | keyof DefaultSchema["Enums"]
    | { schema: keyof DatabaseWithoutInternals },
  EnumName extends DefaultSchemaEnumNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"]
    : never = never,
> = DefaultSchemaEnumNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"][EnumName]
  : DefaultSchemaEnumNameOrOptions extends keyof DefaultSchema["Enums"]
    ? DefaultSchema["Enums"][DefaultSchemaEnumNameOrOptions]
    : never

export type CompositeTypes<
  PublicCompositeTypeNameOrOptions extends
    | keyof DefaultSchema["CompositeTypes"]
    | { schema: keyof DatabaseWithoutInternals },
  CompositeTypeName extends PublicCompositeTypeNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"]
    : never = never,
> = PublicCompositeTypeNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"][CompositeTypeName]
  : PublicCompositeTypeNameOrOptions extends keyof DefaultSchema["CompositeTypes"]
    ? DefaultSchema["CompositeTypes"][PublicCompositeTypeNameOrOptions]
    : never

export const Constants = {
  public: {
    Enums: {
      import_type: [
        "clientes_fornecedores",
        "movimentacoes",
        "vendas",
        "contratos",
      ],
      run_status: ["queued", "processing", "done", "error"],
      upload_status: ["uploaded", "processing", "done", "error"],
    },
  },
} as const
