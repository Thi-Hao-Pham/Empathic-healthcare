# NOTE: Run one case after another
# ============= GP case study  =============
source ("251119. Model_GP.R")

# Extreme test 1: no patients (n_pt_compo1 = 0)
df_inputs_extreme1 <- df_input_base
df_inputs_extreme1$n_pt_compo1 <- 0
df_result_extreme_n_pt0 <- run_cca(df_inputs_extreme1, uptake = 1, burnout = 0)

# Extreme test 2: no eligible subgroups (p_subgroup1 = p_subgroup2 = p_subgroup3 = 0)
df_inputs_extreme2 <- df_input_base
df_inputs_extreme2$p_subgroup1 <- 0
df_inputs_extreme2$p_subgroup2 <- 0
df_inputs_extreme2$p_subgroup3 <- 0
df_result_extreme_p_sub0 <- run_cca(df_inputs_extreme2, uptake = 1, burnout = 0)

# Extreme test 3: no provider (n_providers = 0)
df_inputs_extreme3 <- df_input_base
df_inputs_extreme3$n_providers <- 0
df_result_extreme_n_prov0 <- run_cca(df_inputs_extreme3, uptake = 1, burnout = 0)

# Extreme test 4: very high thresholds
df_inputs_extreme4 <- df_input_base
df_inputs_extreme4$v_threshold_1 <- 10
df_inputs_extreme4$v_threshold_2 <- 100
df_inputs_extreme4$v_threshold_3 <- 100
df_result_extreme_high_thresh <- run_cca(df_inputs_extreme4, uptake = 1, burnout = 0)

# Extreme test 5: very low baseline values (v_mean_1_ctrl = 0.01, v_mean_2_ctrl = 1, v_mean_3_ctrl=1,p_satisfaction_ctrl = 0 
df_inputs_extreme5 <- df_input_base
df_inputs_extreme5$v_mean_1_ctrl <- 0.01
df_inputs_extreme5$v_mean_2_ctrl <- 1
df_inputs_extreme5$v_mean_3_ctrl <- 1
df_inputs_extreme5$p_satisfaction_ctrl <- 0
df_result_extreme_low_base <- run_cca(df_inputs_extreme5, uptake = 1, burnout = 0)

# Extreme test 6: very high trainning cost (c_training_f2f = 1000000)
df_inputs_extreme6 <- df_input_base
df_inputs_extreme6$c_training_f2f <- 1000000
df_result_extreme_high_traincost <- run_cca(df_inputs_extreme6, uptake = 1, burnout = 0)

# Extreme test 7: c_offset_1 = c_offset_2 = c_offset_3 = 1
df_inputs_extreme7 <- df_input_base
df_inputs_extreme7$c_offset_1 <- 1
df_inputs_extreme7$c_offset_2 <- 1
df_inputs_extreme7$c_offset_3 <- 1
df_result_extreme_cost_offset <- run_cca(df_inputs_extreme7, uptake = 1, burnout = 0)

# Extreme test 8: very low effect: smd_1 = smd_2 = smd_3 = 1, smd_sat = -1
df_inputs_extreme8 <- df_input_base
df_inputs_extreme8$smd_1 <- 1
df_inputs_extreme8$smd_2 <- 1
df_inputs_extreme8$smd_3 <- 1
df_inputs_extreme8$smd_sat <- -1
df_result_extreme_low_effect <- run_cca(df_inputs_extreme8, uptake = 1, burnout = 0)


# ---- Export results ----
# export the result in a table with a column "Test" name and the corresponding results
results_list <- list(
  "No patients (n_pt_compo1=0)" = df_result_extreme_n_pt0,
  "No eligible subgroups"        = df_result_extreme_p_sub0,
  "No providers (n_providers=0)" = df_result_extreme_n_prov0,
  "Very high thresholds"         = df_result_extreme_high_thresh,
  "Very low baseline values"     = df_result_extreme_low_base,
  "Very high training costs"     = df_result_extreme_high_traincost,
  "Cost offset = 1"              = df_result_extreme_cost_offset,
  "Very low effect"              = df_result_extreme_low_effect
)

# Convert everything into a 1-row data frame per test
df_extreme_gp <- map_df(names(results_list), function(nm) {
  res <- results_list[[nm]]
  
  # Convert lists to data frames if needed
  res_df <- as.data.frame(t(unlist(res)))
  res_df$Test <- nm
  res_df
})

# Move "Test" column to the front
df_extreme_gp <- df_extreme_gp %>% select(Test, everything(), -strategy)
# ---- Export to EXCEL ----

# Create a new workbook
wb <- createWorkbook()

# ---- Add GP case study results ----
addWorksheet(wb, "extreme_test_GP")
writeData(wb, sheet = "extreme_test_GP", x = df_extreme_gp)

# ---- Save the Excel file ----
saveWorkbook(wb, "Extreme_test_results.xlsx", overwrite = TRUE)

# ============= Recurrent miscarriage case study  =============
source ("251119. Model_miscarriage.R")

# Extreme test 1: no patients (n_pt_compo1 = 0)
df_inputs_extreme1 <- df_input_base
df_inputs_extreme1$n_pt_compo1 <- 0
df_result_extreme_n_pt0 <- run_cca(df_inputs_extreme1, uptake = 1, burnout = 0)

# Extreme test 2: no eligible subgroups (p_subgroup1 = p_subgroup2 = p_subgroup3 = 0)
df_inputs_extreme2 <- df_input_base
df_inputs_extreme2$p_subgroup1 <- 0
df_inputs_extreme2$p_subgroup2 <- 0
df_inputs_extreme2$p_subgroup3 <- 0
df_result_extreme_p_sub0 <- run_cca(df_inputs_extreme2, uptake = 1, burnout = 0)

# Extreme test 3: no provider (n_providers = 0)
df_inputs_extreme3 <- df_input_base
df_inputs_extreme3$n_providers <- 0
df_result_extreme_n_prov0 <- run_cca(df_inputs_extreme3, uptake = 1, burnout = 0)

# Extreme test 4: very high thresholds
df_inputs_extreme4 <- df_input_base
df_inputs_extreme4$v_threshold_1 <- 10
df_inputs_extreme4$v_threshold_2 <- 100
df_inputs_extreme4$v_threshold_3 <- 100
df_result_extreme_high_thresh <- run_cca(df_inputs_extreme4, uptake = 1, burnout = 0)

# Extreme test 5: very low baseline values (v_mean_1_ctrl = 0.01, v_mean_2_ctrl = 1, v_mean_3_ctrl=1,p_satisfaction_ctrl = 0 
df_inputs_extreme5 <- df_input_base
df_inputs_extreme5$v_mean_1_ctrl <- 0.01
df_inputs_extreme5$v_mean_2_ctrl <- 1
df_inputs_extreme5$v_mean_3_ctrl <- 1
df_inputs_extreme5$p_satisfaction_ctrl <- 0
df_result_extreme_low_base <- run_cca(df_inputs_extreme5, uptake = 1, burnout = 0)

# Extreme test 6: very high tranning cost (c_training_f2f = 1000000)
df_inputs_extreme6 <- df_input_base
df_inputs_extreme6$c_training_f2f <- 1000000
df_result_extreme_high_traincost <- run_cca(df_inputs_extreme6, uptake = 1, burnout = 0)

# Extreme test 7: c_offset_1 = c_offset_2 = c_offset_3 = 1
df_inputs_extreme7 <- df_input_base
df_inputs_extreme7$c_offset_1 <- 1
df_inputs_extreme7$c_offset_2 <- 1
df_inputs_extreme7$c_offset_3 <- 1
df_result_extreme_cost_offset <- run_cca(df_inputs_extreme7, uptake = 1, burnout = 0)

# Extreme test 8: very low effect: smd_1 = smd_2 = smd_3 = 1, smd_sat = -1
df_inputs_extreme8 <- df_input_base
df_inputs_extreme8$smd_1 <- 1
df_inputs_extreme8$smd_2 <- 1
df_inputs_extreme8$smd_3 <- 1
df_inputs_extreme8$smd_sat <- -1
df_result_extreme_low_effect <- run_cca(df_inputs_extreme8, uptake = 1, burnout = 0)

# ---- Export results ----
results_list <- list(
  "No patients (n_pt_compo1=0)" = df_result_extreme_n_pt0,
  "No eligible subgroups"        = df_result_extreme_p_sub0,
  "No providers (n_providers=0)" = df_result_extreme_n_prov0,
  "Very high thresholds"         = df_result_extreme_high_thresh,
  "Very low baseline values"     = df_result_extreme_low_base,
  "Very high training costs"     = df_result_extreme_high_traincost,
  "Cost offset = 1"              = df_result_extreme_cost_offset,
  "Very low effect"              = df_result_extreme_low_effect
)

# Convert everything into tidy format
df_extreme_miscarriage <- map_df(names(results_list), function(nm) {
  res <- results_list[[nm]]
  res_df <- as.data.frame(t(unlist(res)))
  res_df$Test <- nm
  res_df
})

df_extreme_miscarriage <- df_extreme_miscarriage %>% select(Test, everything(), -strategy)

# ---- Export to EXCEL ----
# Append miscarriage results to the existing workbook
wb <- loadWorkbook("Extreme_test_results.xlsx")
addWorksheet(wb, "extreme_test_miscarriage")
writeData(wb, sheet = "extreme_test_miscarriage", x = df_extreme_miscarriage)
# ---- Save the Excel file ----
saveWorkbook(wb, "Extreme_test_results.xlsx", overwrite = TRUE)
