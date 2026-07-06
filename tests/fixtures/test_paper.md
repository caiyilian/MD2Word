# Method

![The overview of our proposed SIDA. (a) Cross-Patch Augmentation (CPA) constructs a semantic-preserving intermediate domain by interacting patches between source and target domains, effectively solving the dilemma where weak augmentations lack diversity while strong augmentations destroy semantics. (b) Consistent Prototype Constraint (CPC) aligns global class distributions via prototype contrastive learning and enforces local semantic consistency, addressing the feature distribution discrepancy between domains. (c) Structure-Semantic Pseudo-Labeling and Merging (S2PLM) synergizes the teacher model with MedSAM3D to generate high-quality pseudo-labels with precise boundaries, mitigating the lack of reliable semantic guidance in the unlabeled target domain.](fig/workflow.pdf)

The overview of our proposed SIDA. (a) Cross-Patch Augmentation (CPA) constructs a semantic-preserving intermediate domain by interacting patches between source and target domains, effectively solving the dilemma where weak augmentations lack diversity while strong augmentations destroy semantics. (b) Consistent Prototype Constraint (CPC) aligns global class distributions via prototype contrastive learning and enforces local semantic consistency, addressing the feature distribution discrepancy between domains. (c) Structure-Semantic Pseudo-Labeling and Merging (S2PLM) synergizes the teacher model with MedSAM3D to generate high-quality pseudo-labels with precise boundaries, mitigating the lack of reliable semantic guidance in the unlabeled target domain. (\ref{workflow})

## Preliminary
In UDA settings, given the labeled source data $D_{s}=\{(x^{i}_{s}, y^{i}_{s}) \}_{i=1}^{N_{s}}$ and the unlabeled target domain data $D_{t}=\{x^{i}_{t}\}_{i=1}^{N_{t}}$, our goal is training the model on the labeled source data $D_{s}$ and the unlabeled target data $D_{t}$ to achieve accurate segmentation on the target domain. The data $x_{s}$ and $x_{t} \in \mathbb{R}^{H \times W \times D}$ are 3D images, with the height of $H$, the width of $W$, and the depth of $D$. $y_{s} \in \mathbb{R}^{C \times H \times W \times D}$ denotes the corresponding ground-truth for $C$ classes of the source domain data. 

## Overview
Fig. \ref{workflow} illustrates the framework of the proposed SIDA, an Unsupervised Semantic-preserving Intermediate Domain Adaptation method for medical image segmentation. It consists of three innovative components: Cross-Patch Augmentation (CPA), Structure-Semantic Pseudo-Labeling and Merging (S2PLM), and Consistent Prototype Constraint (CPC).

The overall workflow proceeds as follows. First, taking the labeled source domain data and unlabeled target domain data as inputs, the CPA module constructs a semantic-preserving intermediate domain by interacting patches between the source and target domains. This process enriches data diversity while preserving essential semantic information. Subsequently, the data are fed into a teacher-student framework, where the teacher network generates initial pseudo-labels for the unlabeled data from both the target and intermediate domains. To mitigate the lack of reliable supervision, the S2PLM is employed to merge pseudo-labels derived from different domains, generating high-quality pseudo-labels for training. Finally, the student network is optimized using the ground truth of the source domain and the refined pseudo-labels of the unlabeled data via segmentation losses. Meanwhile, the CPC module imposes a semantic consistency constraint across the source and target domains to align feature distributions, effectively bridging the domain gap and enhancing the model's generalization ability.

## Cross-Patch Augmentation
To effectively bridge the domain gap, we propose the Cross-Patch Augmentation module. CPA is designed to construct a comprehensive intermediate domain $D_m$ that not only covers diverse style distributions but also enforces structural interaction between the source and target domains.
First, to explore the potential style variations within each domain, we employ Shuffle Remap~\cite{kong2023indescribable} as a base augmentation. SR randomly remaps pixel intensity ranges to generate augmented views $x_{sa}$ and $x_{ta}$ from the original inputs $x_s$ and $x_t$, effectively expanding the intra-domain feature space.
Specifically, the distribution of an input $x$ is normalized to $[-1, 1]$ and divided into $N$ segments by randomly generated control points. These segments are then randomly remapped to new ranges, as formulated in Eq.(\ref{eq1}):
\begin{align}  
    x'=\frac{x-P_{i}}{P_{i+1}-P_{i}} \times  (P_{j+1}-P_{j}) + P_{j}.
\end{align}
SR is a completely random remapping method that alters the relative relationships of the overall image distribution. While this capability is beneficial for UDA by significantly enriching data diversity, it inevitably causes severe destruction of semantic information. Relying solely on SR would generate images with distorted structures, potentially misleading the model and hindering the learning of precise semantic boundaries. Furthermore, style augmentation alone does not explicitly model the relationship between the two domains.
To address this, CPA introduces a cross-domain interaction mechanism via a Copy-Paste strategy~\cite{bai2023bidirectional}. By swapping patches between the source and target domains as well as their SR-augmented counterparts, CPA creates hybrid samples that embed the semantic content of one domain into the contextual style of the other. Formally, given $x_s$, $x_t$, $x_{sa}$, and $x_{ta}$, we generate the intermediate domain samples as follows:
\begin{align}
    \begin{split} 
       & x_{t\to s} = x_{t} \odot \mathit{M}_{\alpha} + x_{s} \odot (\mathbf{1} - \mathit{M}_{\alpha}),  \\
       & x_{s\to t} = x_{s} \odot \mathit{M}_{\alpha} + x_{t} \odot (\mathbf{1} - \mathit{M}_{\alpha}), \\
       & x_{t\to sa} = x_{t} \odot \mathit{M}_{\alpha} + x_{sa} \odot (\mathbf{1} - \mathit{M}_{\alpha}),  \\
       & x_{s\to ta} = x_{s} \odot \mathit{M}_{\alpha} + x_{ta} \odot (\mathbf{1} - \mathit{M}_{\alpha}), 
    \end{split} 
\end{align}
where $\odot$ denotes element-wise multiplication, and $\mathit{M}_{\alpha} \in \{0, 1\}^{H \times W \times D}$ is a randomly generated one-centered mask indicating the patch region to be swapped. The notation $x_{A \to B}$ signifies mapping patches from domain A into the context of domain B. For instance, $x_{t \to s}$ represents embedding Target patches into the Source context.
Through this interaction, CPA achieves two key objectives: (1) **Semantic Injection**: $x_{s \to t}$ and $x_{s \to ta}$ inject reliable source semantics along with corresponding ground truth labels into the target-style environment. Although the background remains in the target style (unlabeled), the supervision signal is derived from the ground truth labels of the injected source patches, enabling deterministic supervised training on these specific regions within the target-domain context; (2) **Contextual Adaptation**: $x_{t \to s}$ and $x_{t \to sa}$ place target patterns into the familiar source environment, helping the model adapt to target textures without losing structural guidance. This results in a semantic-preserving intermediate domain $D_m$ that effectively connects the source and target distributions.

## Structure-Semantic Pseudo-Labeling and Merging
To mitigate the lack of semantic guidance in the unlabeled target domain, we propose the Structure-Semantic Pseudo-Labeling and Merging (S2PLM) strategy. Standard pseudo-labeling methods often rely solely on confidence thresholding, which tends to discard boundary regions with lower confidence, leading to fragmented pseudo-labels. To address this, S2PLM leverages the structural universality of MedSAM3D~\cite{wang2025sam} to refine pseudo-labels and utilizes the constructed intermediate domain to enforce a multi-view consistency constraint, thereby generating high-quality pseudo-labels with accurate boundaries and reliable semantics.

First, we employ a teacher-student framework~\cite{tarvainen2017mean, ma2024constructing,zhang2024mapseg} to provide stable predictions. For any unlabeled input $x$ (from $D_{t}$ or $D_{m}$), the teacher model $f_{t}$ generates a probability map $p = f_{t}(x)$. The initial pseudo-label $\hat{y}$ and the confidence mask $m$ are obtained by:
\begin{align} 
    \begin{split} 
       \hat{y} = \mathop{\arg\max}\limits_{c}(p), \quad \hat{m} = \mathbb{I}(\max(p) \ge \tau),
    \end{split} 
\end{align}
where $\tau$ is a confidence threshold, $\mathop{\arg\max}(\cdot)$ extracts the class index, and $\mathbb{I}(\cdot)$ is the indicator function. The teacher model $f_{t}$ is updated by the exponential moving average (EMA) of the student model.

However, the boundary regions of $\hat{y}$ are often filtered out by the threshold $\tau$. To recover these structural details, we introduce a Structure-Semantic Refinement module using MedSAM3D. We utilize the teacher model's high-confidence predictions $\hat{y}^{raw}$ to generate geometric prompts, including point prompts $\mathcal{P}_{pts}$ (centroids) and dense mask prompts. These prompts, along with the original image $x$, are fed into the frozen MedSAM3D to generate the refined pseudo-label $\hat{y}^{ref}$:
\begin{align} 
    \hat{y}^{ref} = \text{MedSAM3D}(x, \mathcal{P}_{pts}, \hat{y}^{raw}) \cup \hat{y}^{raw}.
\end{align}
Here, $\cup$ represents the pixel-wise union operation. This fusion strategy integrates the predictions from both models: regions identified by SAM are assigned high confidence (i.e., setting the corresponding $\hat{m}^{ref}$ to 1), while high-confidence regions originally predicted by the teacher are preserved even if missed by SAM. This ensures a comprehensive merging of reliable semantics from both the teacher's stable predictions and SAM's precise boundaries.

Since the intermediate domain images $x_{t\to s}$ and $x_{s\to t}$ are constructed by mixing source and target patches, we can partially utilize the ground truth labels from the source domain. To maximize supervision accuracy, we compose the labels for these mixed images by combining the source ground truth $y_{s}$ with the refined teacher's predictions for the target regions:
\begin{align}
    \begin{split} 
           & m_{s\to t}  = \hat{m}_{s \to t}^{ref} \oplus \mathit{M}_{\alpha}, \\
           & m_{t\to s}  = \hat{m}_{t \to s}^{ref} \oplus (\mathbf{1} - \mathit{M}_{\alpha}),\\
           & \hat{y}_{s\to t} = y_{s} \odot \mathit{M}_{\alpha} + \hat{y}_{s\to t}^{ref} \odot (\mathbf{1} - \mathit{M}_{\alpha}), \\
           & \hat{y}_{t\to s} = \hat{y}_{t \to s}^{ref} \odot \mathit{M}_{\alpha} + y_{s} \odot (\mathbf{1} - \mathit{M}_{\alpha}),             
    \end{split} 
\end{align}
where $\hat{y}^{ref}$ and $\hat{m}^{ref}$ denote the refined pseudo-label and mask generated by the SAM-based module. The symbol $\oplus$ represents the pixel-wise OR operation. The modified masks $m_{s\to t}$ and $m_{t\to s}$ trust the ground truth regions implicitly (mask set to 1) and use the confidence mask for target regions. To alleviate information destruction introduced by the strong SR augmentation, the data $x_{t\to sa}$ and $x_{s\to ta}$ share the same pseudo-labels as their SR-free counterparts $x_{t\to s}$ and $x_{s\to t}$, respectively.

Standard pseudo-labeling on the target image $x_{t}$ often suffers from noise due to domain shift. To improve reliability, we introduce a consistency check based on the hypothesis that a robust prediction should remain invariant to contextual changes. We observe the predictions for the target regions under two different contexts: directly from the original image $x_{t}$ and from the mixed images $x_{t\to s}$ and $x_{s\to t}$. 
\begin{align} 
    \begin{split} 
      & \hat{y}_{rec} = \hat{y}_{t\to s} \odot \mathit{M}_{\alpha} + \hat{y}_{s\to t} \odot (\mathbf{1} - \mathit{M}_{\alpha}), \\
      & m_{rec} = m_{t\to s} \odot \mathit{M}_{\alpha} + m_{s\to t} \odot (\mathbf{1} - \mathit{M}_{\alpha}).
    \end{split} 
\end{align}
Then, we compare this reconstructed label with the direct prediction $\hat{y}_{dir}$ obtained from $x_{t}$. A pixel is considered reliable only if the model makes consistent predictions across both views and exhibits high confidence. The final mask $m_{t}$ is computed as:
\begin{align} 
      m_{t} = m_{rec} \odot m_{dir} \odot \mathbb{I}(\hat{y}_{dir} = \hat{y}_{rec}),
\end{align}
where $m_{dir}$ is the confidence mask of $x_{t}$. This strict filtering strategy effectively removes unstable predictions prone to context dependency. The final pseudo-label $\hat{y}_{t}$ for $x_{t}$ is set as $\hat{y}_{dir}$.

For labeled data $x_{s}$, we implement the standard dice loss in the training of the student model:
\begin{align} 
    L_{D}(y, \hat{y}) = 1 - 2 \frac{\left \| \hat{y} \cdot y  \right \|_{1}}{\left \| \hat{y}  \right \|_{2}^{2} + \left \| y  \right \|_{2}^{2} }, \\
    L_{Seg}^{s} = \frac{1}{N_{s}} \sum_{i=1}^{N_{s}}(L_{D}(y_{s}^{i}, f(x_{s}^{i}))), 
\end{align}
where $ L_{Seg}^{s}$ denotes the loss of the source domain, $N_{s}$ denotes the amount of source domain data for training, $f(\cdot)$ represents the student model. $L_{D}(\cdot, \cdot)$ denotes the standard dice loss.

For unlabeled data $x_{t}$ and augmented data $x_{m} \in D_{m}$, the pseudo labels and confidence masks generated are used as supervision to guide the training student model in the masked dice loss $L_{D}^{mask}(\cdot, \cdot, \cdot)$, expressed as follows:
\begin{align} 
    \begin{split}
         & L_{Seg}^{m} = \frac{1}{N_{m}} \sum_{i=1}^{N_{m}}(L_{D}^{mask}(\hat{y}_{m}^{i}, f(x_{m}^{i}), m_{m}^{i})),  \\
         & L_{Seg}^{t} = \frac{1}{N_{t}} \sum_{i=1}^{N_{t}}(L_{D}^{mask}(\hat{y}_{t}^{i}, f(x_{t}^{i}), m_{t}^{i})),  \\
         & L_{D}^{mask}(y, \hat{y}, m) = 1 - 2 \frac{\left \| m \cdot \hat{y} \cdot y  \right \|_{1}}{ m \cdot (\left \| \hat{y}  \right \|_{2}^{2} + \left \| y  \right \|_{2}^{2})},
    \end{split}
\end{align}
where $L_{Seg}^{m}$ denotes the loss of the intermediate domain, $L_{Seg}^{t}$ denotes the loss of the target domain, and $N_{m}$, $N_{t}$ denote the amount of data from the intermediate domain and target domain, separately.

## Consistent Prototype Constraint 
To explicitly align feature distributions and enhance class separability, we propose Consistent Prototype Constraint (CPC). While CPA achieves alignment at the pixel level, distribution discrepancies in deep features may persist. CPC addresses this by enforcing two complementary constraints: global distribution alignment via prototypes and local semantic consistency against perturbations.
Specifically, our framework involves 8 inputs: the original source and target images ($x_s, x_t$), their SR-augmented versions ($x_{sa}, x_{ta}$), and the four CP-mixed images ($x_{t \to s}, x_{s \to t}, x_{t \to sa}, x_{s \to ta}$), resulting in 8 corresponding feature maps.

First, to achieve global alignment, we construct a global prototype bank $\mathcal{P} = \{p_1, p_2, ..., p_C\}$ to represent the standard semantic features of each class. We utilize the feature $\boldsymbol{F}_s$ derived from the source domain data $x_s$ with ground truth labels to update these prototypes via a moving average strategy, ensuring they serve as stable anchors for the semantic space.

For the remaining 7 features, we carefully select those suitable for prototype contrastive learning. Since the SR augmentation involves intensive pixel remapping that causes significant feature distribution shifts, the features from SR-related images ($x_{sa}, x_{ta}, x_{t \to sa}, x_{s \to ta}$) may deviate too far from the standard prototypes. Therefore, we only employ the features from the non-SR images—specifically the original target image ($x_t$) and the CP-augmented images ($x_{t \to s}, x_{s \to t}$)—to compute the prototype contrastive loss $L_{pc}$.
We compute the class centroids $c_t$ for these selected features based on the high-confidence pseudo-labels generated by S2PLM. $L_{pc}$ pulls these centroids towards their corresponding source prototypes (positive pairs) while pushing them away from other prototypes (negative pairs), ensuring that features from the same class cluster together across domains. It is formulated as:
\begin{align} 
  L_{pc} = - \frac{1}{K} \sum_{k=1}^{K} \log \frac{\exp(\text{sim}(c_t^k, p_k)/\epsilon)}{\sum_{j=1}^{C} \exp(\text{sim}(c_t^k, p_j)/\epsilon)},
\end{align}
where $K$ is the number of valid classes in the current batch, $\text{sim}(\cdot, \cdot)$ denotes cosine similarity, and $\epsilon=0.1$ is a temperature hyperparameter.

While $L_{pc}$ effectively aligns the global distributions, simply excluding the SR-augmented features would leave them unconstrained. To address this, we incorporate the semantic consistency loss $L_{sc}$ as a complementary constraint. This loss enforces that the semantic features remain invariant to the strong perturbations. Specifically, we maximize the consistency between the features of the SR-free images $\boldsymbol{F}_{d}$ and their SR-augmented versions $\boldsymbol{F}_{da}$:
\begin{align} 
  L_{sc} = - \frac{1}{N_{d}} \sum_{i=1}^{N_{d}} 
  \frac{y_{d}\cdot \boldsymbol{F}_{d} \cdot \boldsymbol{F}_{da}}
  {\left \| y_{d} \cdot \boldsymbol{F}_{d}  \right \|_{2} \cdot \left \| y_{d} \cdot \boldsymbol{F}_{da}  \right \|_{2}},
\end{align}
where $d \in \{s, m, t\}$ denotes the domain. Specifically, $y_s$ is the ground-truth label of the source domain. For the intermediate domain, $y_m$ is composed of $\hat{y}_{t\to s}$ for $x_{t\to s}$ and $x_{t\to sa}$, and $\hat{y}_{s\to t}$ for $x_{s\to t}$ and $x_{s\to ta}$. For the target domain, $y_t = \hat{y}_{rec}$, since the reconstructed label incorporates cross-domain knowledge and provides more robust semantic guidance than the direct prediction.

These two loss functions are mutually reinforcing and indispensable. $L_{pc}$ establishes the correct semantic destinations (prototypes) using the stable features ($x_t, x_{t \to s}, x_{s \to t}$), resolving the domain shift at the category level. Meanwhile, $L_{sc}$ acts as a regularizer, tethering the strongly augmented (SR) features to these stable anchors, ensuring the model's mapping is robust to intense variations. The total feature alignment loss is defined as $L_{feat} = L_{pc} + L_{sc}$. By integrating prototype contrastive learning with augmentation consistency, CPC effectively bridges the domain gap and learns discriminative domain-invariant representations. This process operates as an iterative self-training cycle: improved feature alignment leads to higher quality pseudo labels, which in turn results in more accurate prototypes and centroids.
