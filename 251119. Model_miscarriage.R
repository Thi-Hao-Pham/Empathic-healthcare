########## Miscarriage case ##########
rm(list = ls())
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(purrr)
library(openxlsx)
library(gridExtra)
library(grid)

# ===================== LOAD & CLEAN =====================
df_raw <- read_excel("250827. Inputs.xlsx", sheet = "miscarriage_full") %>%
  filter(For_R == 1) %>%
  select(Input, Value, DSA_low, DSA_high, PSA_dis, PSA_param1, PSA_param2, Label_R) %>%
  mutate(
    Value = as.numeric(
      gsub(",", ".", trimws(as.character(Value))) # replace comma decimal, trim spaces
    ))

df_input_base <- df_raw %>%
  select(Input, Value) %>%
  pivot_wider(names_from = Input, values_from = Value)

df_label <- df_raw %>%
  transmute(
    key = Input,
    label_short = coalesce(na_if(trimws(Label_R), ""), Input)
  )

# ===================== DETECT OUTCOMES DYNAMICALLY =====================
# An outcome i is considered "available" if v_mean_i_ctrl exists; we’ll later
# require (and check) v_mean_i_int, sd_i, v_threshold_i, c_offset_i.
outcome_nums <- df_raw %>%
  filter(grepl("^v_mean_[0-9]+_ctrl$", Input)) %>%
  mutate(num = sub("^v_mean_([0-9]+)_ctrl$", "\\1", Input)) %>%
  pull(num) %>% unique()


# ===================== HELPER FUNCTIONS =====================

## function to calculate the number of patients with improved outcomes
count_below_threshold <- function(n_total, threshold, mean_arm, sd_common, uptake = 1) {
  p_imp <- pnorm(threshold, mean = mean_arm, sd = sd_common)
  ceiling(n_total * p_imp * uptake)
}

# helper function for getting label from df_label
get_label <- function(key, df_label) {
  v <- df_label$label_short[match(key, df_label$key)]
  ifelse(is.na(v) | v == "", key, v)
}


# ============== RUNNER ==============
run_cca <- function(params, uptake = 1, burnout = 0, complaint = 0, format = c("f2f","online")) {
  format <- match.arg(format)
  with(as.list(params), {
    n_patients_eligible <- n_pt_compo1*n_pt_compo2
    n_providers <- (n_pt_compo1*n_pt_compo2)*n_provider_compo1/n_provider_compo2
    
    # pick training config by format
    cfg <- switch(format,
                  "f2f"    = list(cost_train = c_training_f2f,    hours = v_training_hours_f2f,    sessions = n_training_sessions),
                  "online" = list(cost_train = c_training_online, hours = v_training_hours_online, sessions = n_training_sessions)
    )
    
    row <- list(strategy = format)
    total_offsets <- 0
      
      # --- outcomes loop (dynamic) ---
      for (i in outcome_nums) {
        n_tot <- n_patients_eligible
        thr   <- params[[paste0("v_threshold_", i)]]
        m_ctl <- params[[paste0("v_mean_", i, "_ctrl")]]
        sd_i  <- params[[paste0("sd_", i)]]
        smd_i <- params[[paste0("smd_", i)]]
        
        m_int_name <- paste0("v_mean_", i, "_int")
        if (!is.null(params[[m_int_name]]) && !is.na(params[[m_int_name]])) {
          m_int <- params[[m_int_name]]
        } else {
          m_int <- m_ctl + smd_i * sd_i
        }
        
        c_off <- params[[paste0("c_offset_", i)]]
        
        n_int <- count_below_threshold(n_tot, thr, m_int, sd_i, uptake)
        n_ctl <- count_below_threshold(n_tot, thr, m_ctl, sd_i, 1)
        
        delta <- (n_int - n_ctl)
        row[[paste0("imp_outcome", i)]] <- delta
        
        this_offset <- (n_int - n_ctl) * c_off
        row[[paste0("cost_offset_outcome", i)]] <- this_offset
        total_offsets <- total_offsets + this_offset
      }
      
      # --- satisfaction (whole cohort; mixture under partial uptake) ---
      n_sat_ctrl <- p_satisfaction_ctrl * n_patients_eligible
      p_sat_int  <- pnorm(qnorm(p_satisfaction_ctrl) + smd_sat)
      n_sat_int  <- n_patients_eligible * (uptake * p_sat_int + (1 - uptake) * p_satisfaction_ctrl)
      row$imp_satisfied <- round(n_sat_int - n_sat_ctrl)
      
      # --- training & delivering care: COSTS ---
      train_cost <- n_providers * cfg$cost_train * cfg$sessions
      opp_cost   <- n_providers * cfg$hours * c_opp_provider_hourly * cfg$sessions
      visit_add  <- n_patients_eligible * uptake * c_care_component1_per_unit * n_care_component1_units
      gross_cost <- (train_cost + opp_cost + visit_add)
      
      # --- provider TIME (HOURS) ---
      training_hours <- n_providers * cfg$hours * cfg$sessions
      visit_hours    <- (n_patients_eligible * uptake * n_care_component1_units * n_provider_compo1)
      total_provider_hours <- training_hours + visit_hours
      
      # store the components
      row$train_cost <- round(train_cost)
      row$opp_cost   <- round(opp_cost)
      row$visit_add  <- round(visit_add)
      row$gross_cost <- round(gross_cost)
      row$training_hours       <- ceiling(training_hours)
      row$visit_hours          <- ceiling(visit_hours)
      row$total_provider_hours <- ceiling(total_provider_hours)
      
      # --- burnout offsets (optional) ---
      if (burnout == 1) {
        prod_days_avoided <- n_providers * d_absent_per_hcp * r_absent_reduct
        burn_offset <- prod_days_avoided * c_hcp_day_rate
        row$absence_d_avoided <- round(prod_days_avoided)
        row$burn_offset <- round(burn_offset)
      } else {
        burn_offset <- 0
      }
      
      # complaint offsets (optional):
      if (complaint == 1) {
        n_complaints_avoided <- n_complaints * r_complaint_reduct
        complaint_offset <- n_complaints_avoided * c_complaint
        row$n_complaints_avoided <- round(n_complaints_avoided)
        row$complaint_offset <- round(complaint_offset)
      } else {
        complaint_offset <- 0
      }
      
      total_offsets <- total_offsets + burn_offset + complaint_offset
      row$total_cost_offsets <- round(total_offsets)
      
      # --- incremental net cost ---
      incr_cost <- gross_cost - total_offsets
      row$incr_cost <- round(incr_cost)
      
      df_result <- as.data.frame(row, stringsAsFactors = FALSE)
      
      # --- per-1,000 eligible increments (dynamic) ---
      perk <- function(x) 1000 * x / n_patients_eligible
      for (i in outcome_nums) {
        imp_col <- paste0("imp_outcome", i)
        df_result[[paste0("inc_per_1000_outcome", i)]] <- round(perk(df_result[[imp_col]]), 0)
      }
      df_result$inc_per_1000_satisfied <- round(perk(df_result$imp_satisfied), 0)
      row$inc_per_1000_satisfied <- df_result$inc_per_1000_satisfied
      
      # --- incremental net cost per treated patient (€/treated) ---
      denom_treated <- n_patients_eligible * uptake
      df_result$cost_per_patient <- round(df_result$incr_cost / denom_treated,2)
      
      # Order columns
      imp_cols    <- paste0("imp_outcome", outcome_nums)
      inc_cols    <- paste0("inc_per_1000_outcome", outcome_nums)
      offset_cols <- paste0("cost_offset_outcome", outcome_nums)
      common_cols <- c(imp_cols, "imp_satisfied",
                       offset_cols, "total_cost_offsets",
                       "train_cost", "opp_cost", "visit_add", "gross_cost",
                       "incr_cost", "training_hours", "visit_hours", "total_provider_hours",
                       inc_cols,"inc_per_1000_satisfied", "cost_per_patient")
      
      if (burnout == 1) {
        select(df_result, all_of(c(common_cols, "absence_d_avoided", "burn_offset")))
      } else {
        select(df_result, all_of(common_cols))
      }
      if (complaint == 1) {
        select(df_result, everything(), "complaint_offset")
      } else {
        df_result
      }
  })
}

# ============== RUN ==============
# Base case:
df_result_base <- run_cca(df_input_base, uptake = 1, burnout = 0)

# Scenario 1: re-train is needed every 6 months
df_inputs_scenario1 <- df_input_base
df_inputs_scenario1$n_training_sessions <- 2
df_result_scenario1 <- run_cca(df_inputs_scenario1, uptake = 1, burnout = 0)

# Scenario 2: mean outcome under intervention is taken from the trial, not meta-analysis
df_inputs_scenario2 <- df_input_base
df_inputs_scenario2$v_mean_2_int <- 7.77  # from trial
df_result_scenario2 <- run_cca(df_inputs_scenario2, uptake = 1)

# Scenario 3: burnout considered, complaints considered
df_result_scenario3 <- run_cca(df_input_base, uptake = 1, burnout = 1, complaint = 1)

# ================= DSA =================
run_dsa <- function(params, dsa_inputs, uptake = 1, format = c("f2f","online")) {
  format <- match.arg(format)
  results <- list()
  for (i in seq_len(nrow(dsa_inputs))) {
    var <- dsa_inputs$Input[i]
    for (scenario in c("Low", "High")) {
      modified_params <- params
      modified_params[[var]] <- dsa_inputs[[scenario]][i]
      result <- run_cca(modified_params, uptake = uptake, burnout = 0, format = format) %>%
        select(cost_per_patient) %>%
        mutate(
          cost_per_patient = round(cost_per_patient,2),
          dsa_var = var,
          dsa_scenario = tolower(scenario)
        )
      results[[paste(var, scenario, sep = "_")]] <- result
    }
  }
  bind_rows(results)
}

# Load DSA inputs
df_dsa_inputs <- df_raw %>%
  filter(!is.na(DSA_low) & !is.na(DSA_high)) %>%
  select(Input, DSA_low, DSA_high) %>%
  rename(Low = DSA_low, High = DSA_high)

# Run DSA to create df_dsa_results
df_dsa_results <- run_dsa(df_input_base, df_dsa_inputs, uptake = 1, format = "f2f")

######## Tornado diagram ###################
plot_tornado_topn <- function(df_dsa, df_base, format, outcome_var, top_n = 10, title = NULL) {
  base_value <- df_base[[outcome_var]][1]
  
  df_plot <- df_dsa %>%
    select(dsa_var, dsa_scenario, all_of(outcome_var)) %>%
    pivot_wider(names_from = dsa_scenario, values_from = all_of(outcome_var)) %>%
    mutate(
      deviation_low  = low  - base_value,
      deviation_high = high - base_value,
      spread = pmax(abs(deviation_low), abs(deviation_high))
    ) %>%
    arrange(desc(spread)) %>%
    head(top_n) %>%
    mutate(dsa_var = factor(dsa_var, levels = rev(dsa_var)))
  
  df_long <- df_plot %>%
    select(dsa_var, deviation_low, deviation_high) %>%
    pivot_longer(c(deviation_low, deviation_high),
                 names_to = "Scenario", values_to = "Deviation") %>%
    mutate(Scenario = ifelse(Scenario == "deviation_low", "Low", "High"))
  
  p <- ggplot(df_long, aes(x = Deviation, y = dsa_var, fill = Scenario)) +
    geom_bar(stat = "identity", position = "stack") +
    geom_vline(xintercept = 0, linetype = "dashed", color = "black") +
    scale_x_continuous(labels = function(x) x + base_value, name = outcome_var) +
    labs(title = title,
         y = NULL, fill = "Scenario") +
    theme_minimal() +
    theme(panel.grid = element_blank()) +
    theme(panel.grid = element_blank())
  
  p <- p + scale_y_discrete(labels = function(keys) {
    labs <- df_label$label_short[match(keys, df_label$key)]
    ifelse(is.na(labs) | labs == "", keys, labs)  # fallback to the raw key
  })
  p
}

# Plot top 10 variables affecting cost per patient for the f2f format:
tornado_f2f_miscarriage <- plot_tornado_topn(
  df_dsa_results,
  df_result_base,
  format = "f2f",
  outcome_var = "cost_per_patient",
  top_n = 10,
  title = "C"
)
# Export to PNG
ggsave("tornado_f2f_miscarriage.png", plot = tornado_f2f_miscarriage, width = 8, height = 6, dpi = 300)

# ============== PSA ============== 

# Generate PSA input sets
generate_df_input_psa <- function(df_raw, df_input_base, B = 1000, seed = 1) {
  set.seed(seed)
  
  inputs <- names(df_input_base)  # list of all parameter names
  psa_info <- df_raw %>%
    filter(Input %in% inputs) %>%
    select(Input, PSA_dis, PSA_param1, PSA_param2)
  
  # storage
  df_list <- vector("list", B)
  
  for (b in seq_len(B)) {
    draw_vals <- df_input_base  # start from base
    
    for (i in seq_len(nrow(psa_info))) {
      inp <- psa_info$Input[i]
      dis <- psa_info$PSA_dis[i]
      
      if (!is.na(dis) && dis != "") {
        dis <- tolower(trimws(dis))
        p1  <- psa_info$PSA_param1[i]
        p2  <- psa_info$PSA_param2[i]
        
        val <- switch(dis,
                      "beta"     = rbeta(1, shape1 = p1, shape2 = p2),
                      "normal"   = rnorm(1, mean = p1, sd = p2),
                      "lognormal"= rlnorm(1, meanlog = p1, sdlog = p2),
                      "gamma"    = rgamma(1, shape = p1, scale = p2),
                      "uniform"  = runif(1, min = p1, max = p2),
                      stop(sprintf("Unsupported distribution: %s", dis))
        )
        draw_vals[[inp]] <- val
      }
    }
    df_list[[b]] <- draw_vals
  }
  
  # bind into a big wide data frame: B rows x inputs columns
  bind_rows(df_list)
}

df_input_psa <- generate_df_input_psa(df_raw, df_input_base, B = 1000, seed = 1)

# ---------------------------
run_psa <- function(df_input_psa, uptake = 1, format = c("f2f","online")) {
  format <- match.arg(format)
  results <- vector("list", nrow(df_input_psa))
  for (b in seq_len(nrow(df_input_psa))) {
    params <- df_input_psa[b, , drop = FALSE]
    out_b  <- run_cca(params, uptake = uptake, burnout = 0, format = format)
    out_b$draw <- b
    results[[b]] <- out_b
  }
  bind_rows(results)
}

# Run PSA
df_result_psa <- run_psa(df_input_psa, uptake = 1, format = "f2f")

prob_cost_saving <- df_result_psa %>%
  summarise(
    draws = n(),
    mean_cost = mean(cost_per_patient, na.rm = TRUE),
    median_cost = median(cost_per_patient, na.rm = TRUE),
    p_cost_saving = mean(cost_per_patient < 0, na.rm = TRUE)
  ) %>%
  arrange(desc(p_cost_saving))

print(prob_cost_saving)

# --- 2) Violin plot of incremental cost per patient ---
violin_plot_psa_miscarriage <- ggplot(df_result_psa, aes(x = "", y = cost_per_patient)) +
  geom_violin(fill = "#00BFC4", color = "black", width = 0.5, trim = FALSE, alpha = 0.7) +
  geom_boxplot(width = 0.1, outlier.alpha = 0.2, fill = "white") +
  geom_hline(yintercept = 0, linetype = 2) +
  labs(
    x = NULL,
    y = "Incremental cost per patient (£)",
    title = "D",
  ) +
  theme_minimal() +
  theme(legend.position = "none") +
  theme(panel.grid = element_blank())

sen_plots <- grid.arrange(
  tornado_f2f_miscarriage, nullGrob(), violin_plot_psa_miscarriage,
  ncol = 3, widths = c(2.5, 0.2, 1))

# Export to PNG
ggsave("violin_plot_psa_miscarriage.png", plot = violin_plot_psa_miscarriage, width = 8, height = 6, dpi = 300)
ggsave("sen_plots_miscarriage.png", plot = sen_plots, width = 12, height = 6, dpi = 300)
# ================= EXPORT =================
# Load the exporter
source("export_helpers.R")

# Build the scenario list for this case
res_list <- list(
  "Base-case"  = df_result_base,
  "Scenario 1" = df_result_scenario1,
  "Scenario 2" = df_result_scenario2,
  "Scenario 3" = df_result_scenario3
)

# Run the export 
export_case_study(
  res_list      = res_list,
  outcome_nums  = outcome_nums,
  out_file      = "251119. CCA_results_miscarriage_case.xlsx",   
  tornado_plot  = tornado_f2f_miscarriage,                         
  violin_plot   = violin_plot_psa_miscarriage,                  
  psa_table     = prob_cost_saving,                    
  rename_labels = FALSE                                  
)

########### Stacked bar chart of cost component ##############
## function to label results with descriptive names
label_results <- function(df) {
  lookup <- c(
    "Strategy" = "strategy",
    "Offset cost on stress" = "cost_offset_outcome1",
    "Offset cost on depression" = "cost_offset_outcome2",
    "Training cost" = "train_cost",
    "Opportunity cost" = "opp_cost",
    "Additional consultation cost" = "visit_add",
    "Offset cost on burnout" = "burn_offset",
    "Offset cost on complaints" = "complaint_offset"
    
  )
  # Only keep names that exist in df
  lookup <- lookup[lookup %in% names(df)]
  rename(df, !!!lookup)
}

# create a df_cost_component to store all cost component, turn offset cost to negative and relabel 
df_cost_component <- bind_rows(res_list, .id = "Analysis") %>%
  select(Analysis, train_cost, opp_cost, visit_add, cost_offset_outcome2, burn_offset, complaint_offset) %>%
  label_results() %>%
  mutate(across(-Analysis, ~ . / 1e6),          # convert to £ million
         across(starts_with("Offset"), ~ -abs(.))) %>%  # offsets negative
  pivot_longer(-Analysis, names_to = "Category", values_to = "Value")

ggplot(df_cost_component, aes(x = Analysis, y = Value, fill = Category)) +
  geom_col() +
  labs(title = "",
       y = "Cost (Million £)", x = "") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
  theme(panel.grid = element_blank())

