import numpy as np
import matplotlib.pyplot as plt

v_base = 2.0                      # Base Case Superficial gas velocity in m/s
p_fan_base = 110                  # Base case Fan pressure in kpa

absorber_capex_base = 1.03e7      # Base case Absorber CAPEX in euro
absorber_area_base = 39.18        # Base case Absorber area in m^2

fan_capex_base = 6.877e5          # Base case Fan CAPEX in euro 
fan_capacity_base = 190800        # Base case Fan capacity in m^3/h

fan_opex_base = 268600            # Base case Fan OPEX in euro/year

scaling_exponent = 0.6            # Scaling exponent
ANNUALIZATION_FACTOR_AF = 11.15   # Annualized factor
n = 2                             # Taking n as 2

# Superficial gas velocity range for optimization
velocities = np.arange(1.5, 3.0 + 0.5, 0.5)

# Formula for Absorber area
absorber_area = absorber_area_base * (v_base / velocities)

# Power law for Absorber CAPEX
absorber_capex = absorber_capex_base * (absorber_area / absorber_area_base) ** scaling_exponent

# Formula for Fan capacity
p_fan = p_fan_base * (velocities / v_base) ** n

# Power law for Fan CAPEX
fan_capex = fan_capex_base * (velocities / v_base) ** scaling_exponent

# Power law for Fan OPEX
fan_opex = fan_opex_base * (velocities / v_base) ** 2

# Annualized CAPEX  (€/year)
abs_capex_ann   = absorber_capex / ANNUALIZATION_FACTOR_AF
fan_capex_ann   = fan_capex      / ANNUALIZATION_FACTOR_AF

# Total annualized cost (€/year): Annualized CAPEX + Annual OPEX
total_cost = abs_capex_ann + fan_capex_ann + fan_opex

# Find optimum velocity
opt_index      = np.argmin(total_cost)
opt_velocity   = velocities[opt_index]
opt_total_cost = total_cost[opt_index]

# Print results
print("Gas velocity optimization results (annualized costs):\n")
print(f"{'Velocity (m/s)':>15} {'Area (m^2)':>15} "
      f"{'Abs CAPEX (€/yr)':>18} {'Fan CAPEX (€/yr)':>18} "
      f"{'Fan OPEX (€/yr)':>18} {'Total (€/yr)':>16}")

for i in range(len(velocities)):
    print(f"{velocities[i]:15.2f} "
          f"{absorber_area[i]:15.2f} "
          f"{abs_capex_ann[i]:18.2f} "
          f"{fan_capex_ann[i]:18.2f} "
          f"{fan_opex[i]:18.2f} "
          f"{total_cost[i]:16.2f}")

print("\nOptimum velocity based on minimum annualized total cost:")
print(f"Velocity = {opt_velocity:.2f} m/s")
print(f"Minimum annualized total cost = {opt_total_cost:,.2f} €/yr")

# Stcked bar chart for annualized costs
import matplotlib as mpl
mpl.rcParams['font.family'] = 'Times New Roman'
mpl.rcParams['font.size'] = 13

colors = ['#5DADE2', '#E67E22', '#27AE60']
x = np.arange(len(velocities))
vel_labels = [f"{v:.1f} m/s" for v in velocities]

fig, ax = plt.subplots(figsize=(11, 7))

bars1 = ax.bar(x, abs_capex_ann, color=colors[0], label='Annualized Absorber CAPEX', edgecolor='white', linewidth=0.6)
bars2 = ax.bar(x, fan_capex_ann, color=colors[1], label='Annualized Fan CAPEX',      edgecolor='white', linewidth=0.6, bottom=abs_capex_ann)
bars3 = ax.bar(x, fan_opex,      color=colors[2], label='Annual Fan OPEX',           edgecolor='white', linewidth=0.6, bottom=abs_capex_ann + fan_capex_ann)

# Label values inside each colored segment
for i in range(len(velocities)):
    # Absorber CAPEX — centred inside its segment
    mid1 = abs_capex_ann[i] / 2
    ax.text(x[i], mid1, f'{abs_capex_ann[i]/1e6:.3f} M€',
            ha='center', va='center', fontsize=12, color='white',
            fontfamily='Times New Roman', fontweight='bold')
    # Fan CAPEX — centred inside its segment
    mid2 = abs_capex_ann[i] + fan_capex_ann[i] / 2
    ax.text(x[i], mid2, f'{fan_capex_ann[i]/1e6:.3f} M€',
            ha='center', va='center', fontsize=12, color='white',
            fontfamily='Times New Roman', fontweight='bold')
    # Fan OPEX — centred inside its segment
    mid3 = abs_capex_ann[i] + fan_capex_ann[i] + fan_opex[i] / 2
    ax.text(x[i], mid3, f'{fan_opex[i]/1e6:.3f} M€',
            ha='center', va='center', fontsize=12, color='white',
            fontfamily='Times New Roman', fontweight='bold')

# Total annualized cost on top of each bar
for i, total in enumerate(total_cost):
    ax.text(x[i], total + 0.005e6, f'{total/1e6:.3f} M€/yr',
            ha='center', va='bottom', fontsize=13,
            fontfamily='Times New Roman', fontweight='bold', color='black')

# Horizontal dotted line from Y-axis at the optimum (minimum) cost
ax.axhline(y=opt_total_cost, color='black', linestyle=':', linewidth=1.8,
           label=f'Optimum: {opt_total_cost/1e6:.3f} M€/yr @ {opt_velocity:.1f} m/s')

ax.set_xticks(x)
ax.set_xticklabels(vel_labels, fontsize=13)
ax.tick_params(axis='x', labelsize=13)
ax.tick_params(axis='y', labelsize=13)
ax.set_xlabel("Superficial Gas Velocity in Absorber column (m/s)", fontsize=15, fontweight='bold')
ax.set_ylabel("Annualized Total Cost (Million €/year)", fontsize=15, fontweight='bold')
    # Title removed per user request
ax.legend(loc='lower left', fontsize=12, framealpha=0.9)
ax.yaxis.set_major_formatter(mpl.ticker.FuncFormatter(lambda v, _: f'{v/1e6:.2f} M€'))
ax.yaxis.set_major_locator(mpl.ticker.MultipleLocator(0.2e6)) # 0.2 M€ step increments
ax.set_ylim(bottom=0)
ax.grid(True, axis='y', alpha=0.3, linestyle='--')
fig.tight_layout()
plt.savefig("gas_velocity_optimization.pdf", bbox_inches='tight')

#Total cost plot 
plt.figure(figsize=(8, 5))
plt.bar([f"{v:.1f}" for v in velocities], total_cost)
plt.xlabel("Gas velocity in absorber (m/s)")
plt.ylabel("Total cost (€)")
    # Title removed per user request

#Optimum dotted line
plt.axhline(opt_total_cost, linestyle='--', linewidth=1)
plt.tight_layout()
plt.savefig("total_cost_vs_gas_velocity.pdf", bbox_inches='tight')
