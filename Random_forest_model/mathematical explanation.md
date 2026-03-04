# Mathematical Formulation and Model Interpretation

## 10. Mathematical Foundations of the Models

### 10.1 Random Forest — Formal Definition

Random Forest is based on an ensemble of decision trees trained using bootstrap aggregation.

Given a dataset:

$$\mathcal{D} = \{(x_i, y_i)\}_{i=1}^{N}$$

where:

- $x_i \in \mathbb{R}^d$ is the feature vector of a row
- $y_i \in \{0, 1\}$ indicates whether the row is a header

Each tree $T_b$ is trained on a bootstrap sample $\mathcal{D}_b \subset \mathcal{D}$.

The prediction of a Random Forest with $B$ trees is:

$$\hat{f}(x) = \frac{1}{B} \sum_{b=1}^{B} T_b(x)$$

For classification, the predicted probability is the average of individual tree probabilities:

$$P(y=1 \mid x) = \frac{1}{B} \sum_{b=1}^{B} P_b(y=1 \mid x)$$

The final classification is:

$$\hat{y} = \begin{cases} 1 & \text{if } P(y=1 \mid x) &gt; \tau \\ 0 & \text{otherwise} \end{cases}$$

where $\tau$ is the decision threshold.

Random Forest primarily reduces variance, since averaging independent trees decreases sensitivity to noise. However, it does not iteratively correct errors, which limits its ranking sharpness in highly imbalanced problems.

---

### 10.2 Gradient Boosting — LightGBM Mathematical Formulation

LightGBM is based on Gradient Boosting Decision Trees (GBDT). Unlike Random Forest, trees are built sequentially.

The model is defined as an additive function:

$$F_M(x) = \sum_{m=1}^{M} \gamma_m h_m(x)$$

where:

- $h_m(x)$ is the tree added at iteration $m$
- $\gamma_m$ is its weight
- $M$ is the total number of boosting iterations

The objective is to minimize a differentiable loss function:

$$\mathcal{L} = \sum_{i=1}^{N} l(y_i, F(x_i))$$

For binary classification, the logistic loss is used:

$$l(y, F(x)) = \log(1 + e^{-yF(x)})$$

At each iteration $m$, the model fits a new tree to the negative gradient (pseudo-residuals):

$$r_{im} = -\left[ \frac{\partial l(y_i, F(x_i))}{\partial F(x_i)} \right]$$

Each new tree is trained to approximate these residuals:

$$h_m(x) \approx r_{im}$$

Then the model is updated:

$$F_m(x) = F_{m-1}(x) + \eta h_m(x)$$

where:

- $\eta$ is the learning rate
- $F_{m-1}(x)$ is the previous ensemble

For binary classification, the predicted probability is:

$$P(y=1 \mid x) = \frac{1}{1 + e^{-F_M(x)}}$$

#### Why This Matters for Header Detection

Because each new tree corrects previous errors:

- Difficult header rows receive increasing focus
- Probability separation becomes sharper
- Ranking quality improves

This property is essential in your task, where the final prediction is:

$$\hat{y}_{\text{sheet}} = \arg\max_{\text{row} \in \text{sheet}} P(y=1 \mid x_{\text{row}})$$

Thus, LightGBM is structurally better suited to ranking problems within grouped data.

---

## 11. Feature Importance Analysis — LightGBM Model

Feature importance in LightGBM is computed based on the total contribution of each feature to tree splits across all boosting iterations.

Formally, if a feature $j$ is used in a split that produces information gain $\Delta \mathcal{L}$, the total importance is:

$$\text{Importance}(j) = \sum_{\text{splits on } j} \Delta \mathcal{L}$$

This measures how much each feature reduces the loss function.

### 11.1 Most Influential Features

In your LightGBM implementation, the most important features were:

#### 1️⃣ `num_to_str_ratio`

This feature captures the difference between numeric and textual content within a row:

$$\text{num\_to\_str\_ratio} = \text{num\_ratio} - \text{str\_ratio}$$

Headers typically contain mostly strings, while data rows contain more numeric values. This makes the feature highly discriminative.

#### 2️⃣ `bold_ratio`

Headers often use formatting emphasis. The model strongly relies on this signal when available. This confirms that Excel styling contains predictive information.

#### 3️⃣ `str_ratio` and `num_ratio`

These two foundational features encode the core structural difference between header and data rows. Their combined effect reinforces decision boundaries.

#### 4️⃣ `row_position`

Headers are usually located near the top of the sheet. The normalized position:

$$\text{row\_position} = \frac{\text{row\_index}}{\text{max\_rows}}$$

acts as a positional prior.

#### 5️⃣ Delta Features (`delta_str_ratio`, `delta_num_ratio`)

These contextual features measure contrast with the next row:

$$\Delta_{\text{str}} = \text{str}_{\text{current}} - \text{str}_{\text{next}}$$

Headers are often followed by numeric data, making this difference large and highly informative.

### 11.2 Interpretation

The importance distribution confirms that the model relies on:

- **Structural composition** (text vs numeric)
- **Formatting signals**
- **Local contrast** between adjacent rows
- **Positional bias**

This demonstrates that the feature engineering strategy successfully encodes the structural properties that differentiate headers from regular rows.

Notably, contextual and ratio-based features dominate over raw textual statistics (e.g., string length), indicating that macro-structure matters more than micro-text properties.

---

## 12. Final Analytical Insight

From a modeling perspective:

| Aspect | Random Forest | LightGBM |
|--------|--------------|----------|
| **Mechanism** | Reduces variance via averaging | Reduces bias through sequential error correction |
| **Optimization** | Parallel tree building | Additive gradient boosting |
| **Ranking Quality** | Limited sharpness | Improved probability separation |

Your task benefits from improved ranking sharpness rather than simple classification accuracy.

Mathematically and empirically, LightGBM is better aligned with the constrained per-sheet maximum probability selection strategy.

The combination of:

$$\text{Boosting} + \text{Group-based Ranking} + \text{Structural Feature Engineering}$$

produces a robust and production-ready header detection pipeline.
