# Random_forest_header_detection
Creation, training of an random forest model to detect the header inside an excel document and comparaison with an Heuristic Scoring Algorithm model to compare which one is more efficient for this task


# Header Row Detection in Heterogeneous Excel Files

## Comparative Study Between Random Forest and LightGBM Models

The objective of this project is to design a machine learning system capable of automatically detecting the header row within Excel sheets. The problem is formulated as a supervised binary classification task at the row level: for each row in a sheet, the model predicts whether it corresponds to the true header.

Because each sheet contains exactly one header row and multiple non-header rows, the task presents a strong class imbalance and requires not only accurate classification but also reliable probability ranking.

## Part I — Random Forest Model

###  The Random Forest Approach

The first implementation relied on a Random Forest classifier. Random Forest is an ensemble method based on decision trees trained using bootstrap aggregation (bagging). Each tree is built independently on a random subset of data and features. The final prediction is obtained by averaging predictions across all trees.

The advantage of Random Forest lies in its robustness. It handles non-linear relationships naturally, requires limited preprocessing, and is resistant to overfitting due to feature randomness and bootstrapping. As a baseline model for structured tabular data, it is a logical first choice.


###  Random Forest Parameters

In the initial notebook implementation, the Random Forest model was configured using parameters controlling tree complexity and generalization. The number of estimators determined how many trees were built. The maximum depth constrained how complex each tree could become. The minimum samples per split and class weighting were used to address the strong imbalance between header and non-header rows.

Because each sheet contains only one positive sample and many negatives, the imbalance ratio is structurally high. While Random Forest can partially compensate for imbalance using class weights, it does not inherently optimize probability ranking.

###  Observed Limitations

Although the Random Forest model provided meaningful feature importance insights, it did not perform reliably in production scenarios.

The primary limitation was related to probability calibration and ranking quality. In this project, predictions are not made independently per row. Instead, the row with the highest predicted probability within each sheet is selected as the potential header. Therefore, relative ranking between rows is more important than absolute classification accuracy.

Random Forest averages independent trees and tends to produce less sharply differentiated probability distributions. As a result, several sheets exhibited either false positives (non-header rows with slightly higher probability) or false negatives (true headers not sufficiently separated from competitors).

Additionally, because trees are built independently, Random Forest does not iteratively correct previous mistakes. This limits its ability to refine decision boundaries in highly structured ranking problems.

For these reasons, the Random Forest approach did not meet performance expectations in terms of stability and recall.

## Part II — LightGBM Model

###  Introduction to LightGBM

The second model replaced Random Forest with LightGBM, a gradient boosting framework optimized for efficiency and performance on tabular data.

Unlike Random Forest, which builds trees independently, LightGBM constructs trees sequentially. Each new tree is trained to correct the residual errors of the previous ensemble. This boosting mechanism reduces bias and progressively sharpens the decision boundary.


### Conceptual Differences with Random Forest

The fundamental difference lies in the learning strategy. Random Forest reduces variance by averaging independent trees, while LightGBM reduces both bias and variance by iteratively minimizing prediction errors.

In ranking-sensitive problems such as header detection, boosting provides a structural advantage. Because each new tree focuses on previously misclassified samples, the model becomes increasingly sensitive to subtle structural differences between header and non-header rows.

Furthermore, LightGBM is optimized for speed and memory efficiency, using histogram-based splitting and leaf-wise tree growth. This makes it both faster and often more accurate than traditional ensemble methods.

### LightGBM Configuration

The implemented LightGBM model was configured with 400 estimators and a learning rate of 0.05. This relatively low learning rate ensures gradual refinement of predictions while maintaining stability.

The parameter num_leaves was set to 20, controlling the complexity of individual trees. The maximum depth was limited to 6 to prevent overfitting. The min_child_samples parameter ensured that leaf nodes contained sufficient data points to generalize well.

Subsampling (subsample = 0.8) and feature sampling (colsample_bytree = 0.8) were introduced to reduce variance and improve robustness. A fixed random state ensured reproducibility.

Most importantly, the post-processing strategy was redesigned. Instead of applying a global classification threshold, predictions were grouped by sheet, and only the row with the highest probability per sheet was considered eligible as a header. A threshold of 0.35 was applied to avoid weak predictions.

This transformation effectively converts the task into a constrained ranking problem, significantly improving stability.

### Performance and Final Assessment

Compared to Random Forest, the LightGBM model demonstrated superior ranking behavior and improved balance between precision and recall. False negatives were reduced, and probability separation between header and non-header rows became more pronounced.

The boosting mechanism proved particularly effective in capturing structural contrasts encoded in the delta features. Because each tree corrects prior residual errors, the model progressively emphasizes subtle but discriminative signals.

In conclusion, while Random Forest provided a useful baseline and feature validation framework, LightGBM offered structurally superior performance for this structured tabular detection problem. The sequential boosting process, combined with group-based probability selection, resulted in a more robust and production-ready solution.
