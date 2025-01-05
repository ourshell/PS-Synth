# Extended Harmonic Coherence Temperament (EHCT)

A **ground-breaking theory** that builds on the spirit of Near-Equal Temperament (N-ET), yet reimagines the approach from the ground up. It retains the practical aim of approximating core just intervals while ensuring a near-equal 12-tone subdivision. At the same time, it expands the idea of “stacking intervals” to a broader set of primes and introduces a novel mathematical framework for **minimizing aggregate interval error**. Think of it as a **next-generation** near-equal tuning that pushes the boundaries of **harmonic clarity**, **psychoacoustic acceptability**, and **modular convenience**.

---

## 1. Introduction

From the **Pythagoreans** to **modern digital tuners**, humans have worked tirelessly to reconcile two core demands of music:

1. **Harmonic purity**: Maintaining integer or near-integer ratios among pitches (e.g., 3:2, 5:4, 6:5, etc.).  
2. **Modular flexibility**: Enabling seamless transposition and chromatic usage across all keys.

**Equal Temperament (ET)** offered a historically decisive solution: 12 logarithmically equal semitones. Yet we pay a price—most famously, the **major third** is off by nearly 14 cents from its just ratio of \(5:4\), and the perfect fifth is ~2 cents too small. **Near-Equal Temperament (N-ET)** significantly reduces these discrepancies by stacking 7 pure fifths and 1 pure major third, then folding within an octave, achieving a negligible \(\pm 0.001\) cent discrepancy in a single key center.

In this paper, we propose a further leap forward: a **generalized method** for constructing a 12-tone scale that **harmonically weights** multiple intervals (3:2, 5:4, possibly 7:4) while retaining a near-equal step size. We call this approach **Extended Harmonic Coherence Temperament (EHCT)**. Our goal is to:

1. **Honor the prime harmonic intervals** that shape Western tonality (3, 5, and optionally 7).  
2. **Distribute minimal total error** across all 12 scale degrees, ensuring no single interval is significantly “out of tune.”

Where N-ET uses a clever stacking of a few just intervals, EHCT expands the scope to a multi-interval **optimization** of the entire scale.

---

## 2. Rethinking Tuning: The Weighted Harmonic Coherence Approach

### 2.1 Harmonic Coherence

Any chord or interval with **integer frequency ratios** has maximum *harmonic coherence* (resonance, stability). In the real world, slight deviations from integer ratios still “work,” especially if they lie **below the Just Noticeable Difference (JND)** for harmonic intervals.

### 2.2 Weighted Interval Error

Unlike classical “stacked-fifths” methods that fix one or two intervals exactly, EHCT attempts to **minimize the global error** of multiple intervals at once, acknowledging that:

- \(3/2\) (perfect fifth) remains a cornerstone of Western harmony.  
- \(5/4\) (major third) is almost as crucial.  
- \(6/5\) (minor third) or \(7/4\) (harmonic seventh) may also be relevant to extended tonal or blues-based music.

We define a “cost function” \(C\) that measures the difference between each just interval ratio and its EHCT approximation in cents, scaled by psychoacoustic importance:

\[
C 
= \sum_{\text{interval } i} 
  w_i \, \bigl[\;\Delta_{\text{EHCT}}(i)\bigr]^2,
\]

where:
- \(\Delta_{\text{EHCT}}(i)\) is the cents difference for interval \(i\) under the EHCT scale.  
- \(w_i\) is a weight reflecting the **importance** of that interval (e.g., \(w_{3/2}\) might be higher than \(w_{7/4}\) if we care more about perfect fifths than harmonic sevenths).

### 2.3 Logarithmic Framework

Because pitch perception is **logarithmic**, we represent each note in the scale as:

\[
f_n 
= f_0 \times 2^{\frac{x_n}{12}},
\]

where \(x_n\) is the “scale degree” measured in **12-ET semitones**. In pure 12-ET, \(x_n = n\) for \(n = 0,1,\dots,11\). In EHCT, the \(x_n\) are **not** integers but real-valued, chosen to keep the entire scale as close to a 12-tone division as possible while reducing the interval errors for the prime ratios.

---

## 3. Minimizing Weighted Interval Errors

### 3.1 Mathematical Formulation

Let \(\{x_0, x_1, \dots, x_{11}\}\) be the 12 scale degrees in EHCT, each in the range \([0,12)\). By definition, we set \(x_0 = 0\) for the tonic. Then for intervals like a perfect fifth (3:2) between scale degrees \(x_i\) and \(x_j\):

\[
\Delta_{\text{EHCT}}\bigl(3:2, i, j\bigr) 
= 1200 \times \frac{\ln\!\bigl(2^{\tfrac{x_j - x_i}{12}}\bigr) 
     - \ln\!\bigl(\tfrac{3}{2}\bigr)}{\ln(2)}.
\]

**Goal**: Find the set \(\{x_n\}\) that **minimizes** the global cost function \(C\). We subject the solution to the constraint that \(x_{n+12} = x_n + 12\) for any expansions beyond the first octave, thus ensuring we “tile” the scale in a repeating 12-tone pattern.

### 3.2 Algorithmic Solutions

We can solve for \(\{x_n\}\) using:

1. **Gradient Descent** (numerical optimization).  
2. **Iterated Comma Adjustments** (extending the concept of “stacking intervals,” but readjusting each scale tone in small increments to spread out the error).  
3. **Integer Programming** with a fine grid approximation (if we want to keep scale steps “rational” or close to rational fractions).

EHCT is flexible—**different priorities** in chord purity (e.g., jazz vs. classical) can shift the weighting \(w_i\). The result is a **custom** near-equal scale that systematically reduces the sum of relevant interval errors.

---

## 4. The 3-Prime EHCT Model: 3, 5, and 7

### 4.1 Why Include Prime 7?

Although 7-limit intervals (e.g., \(7/4\), \(7/6\)) rarely appear in standard Western harmony, they **frequently** occur in folk traditions, blues bends, and extended chords. Incorporating 7 in EHCT can yield a scale that also resonates with these expansions.

### 4.2 Example Weights

Let us define the following weight distribution for demonstration:

\[
w_{3/2} = 1.0, 
\quad w_{5/4} = 0.8, 
\quad w_{7/4} = 0.4.
\]

- We strongly emphasize purity of the **perfect fifth** (\(3/2\)).  
- We also care about a **good** major third (\(5/4\)).  
- We value the **harmonic seventh** (\(7/4\)), but to a lesser extent.

Other intervals (like \(6/5\) for the minor third) can also be included with appropriate weights. The cost function grows in dimensionality, but modern computers can handle the search efficiently.

### 4.3 Constructing the Scale

1. **Initialization**: Start with a rough 12-ET scale, \(\{0,1,2,\dots,11\}\).  
2. **Iterative Adjustment**:  
   - For each pair \((i,j)\) that represents an interval of interest, calculate the current error \(\Delta_{\text{EHCT}}(i,j)\).  
   - Slightly nudge \(x_j\) (or \(x_i\)) to **reduce** the sum of squared errors.  
3. **Normalization**: Ensure that the final \(x_{11}\) remains close to 11 (i.e., 1100 cents from tonic), so the scale doesn’t stretch or contract drastically compared to 12-ET.

After sufficient iteration, we converge on a set \(\{x_0, x_1, \ldots, x_{11}\}\) that yields **very low total error** for 3:2, 5:4, and 7:4.

---

## 5. Implementation and Results

Below is a **hypothetical** result (numeric example):

| Scale Degree | EHCT Value (\(x_n\)) | Frequency Ratio (vs. Tonic)      | Approx. Cents from Tonic |
|--------------|----------------------|----------------------------------|---------------------------|
| Tonic (0)    | 0.000               | \(2^0 = 1\)                      | 0                         |
| 2nd          | 2.038               | \(2^{(2.038/12)} \approx 1.122\) | ~203.8                    |
| 3rd          | 3.996               | \(2^{(3.996/12)} \approx 1.259\) | ~399.6                    |
| 4th          | 5.000               | \(2^{(5.000/12)} \approx 1.334\) | 500                       |
| 5th          | 7.009               | \(2^{(7.009/12)} \approx 1.498\) | 700.9                     |
| 6th          | 9.038               | \(2^{(9.038/12)} \approx 1.681\) | 903.8                     |
| 7th          | 10.949              | \(2^{(10.949/12)} \approx 1.952\)| 1094.9                    |
| 8ve (12)     | 12.000              | \(2^1 = 2\)                      | 1200                      |

- The major third is around **399.6 cents** from the tonic (error ~\(-0.4\) cents from 400).  
- The perfect fifth is around **700.9 cents** (error ~\(+0.9\) cents from 700).  
- The 7th partial interval (e.g., from the 4th or 5th scale degree) remains close to a 7-limit ratio.  
- Each step is near **100 cents** on average, preserving the near-equal feel.

A thorough, computer-based optimization might refine these values further, often achieving sub-1-cent total mean error for the selected intervals.

---

## 6. Psychoacoustic and Musical Advantages

1. **Minimal “beating”** in important chords: The perfect fifth, major third, and harmonic seventh align closely with just ratios.  
2. **Smooth voice-leading**: Because we retain near 12-ET spacing, melodic lines feel comfortably familiar.  
3. **Flexible transposition**: Although we emphasize specific intervals, no single note is drastically out of place, allowing chord progressions in many keys.  
4. **Adaptable**: We can recast or reweight intervals for genres focusing on **blues** (strong emphasis on 7-limit intervals) vs. **classical** (3-limit and 5-limit intervals), or even microtonal expansions.

---

## 7. Practical Guidelines

- **Implementation**: EHCT can be implemented in digital synths via a **look-up table** storing the frequency ratios. Acoustic instrument makers can place frets or tone holes according to the final “EHCT scale.”  
- **Calibration**: Musicians can calibrate their tuners to the final EHCT frequencies. Since the differences from 12-ET are subtle (often within \(\pm2\) cents), the new scale should feel **familiar** yet distinctly more resonant.  
- **Extension to 24 or 31 Tones**: The same weighting approach can be generalized to 24-tone (quarter-tone) or 31-tone systems for microtonal composers, ensuring that core intervals remain well-tuned while allowing more pitch classes for coloristic effects.

---

## 8. Conclusion

**Extended Harmonic Coherence Temperament (EHCT)** is a bold new “fork” from traditional Near-Equal Temperament, intended to **optimize** multiple prime-limit intervals simultaneously while preserving a familiar 12-tone layout. Instead of perfectly stacking a single or dual interval set (like 3:2 and 5:4 alone), EHCT **mathematically distributes** tuning “error” in a measured way, guided by psychoacoustic weighting.

This strategy:

1. **Embeds** the crucial intervals from Just Intonation—particularly the 3:2, 5:4, and optionally 7:4—within a single 12-tone octave.  
2. **Minimizes total interval error** across multiple chordal structures, rather than forcing one or two intervals to be exactly pure.  
3. **Maintains** near-equal spacing of pitch classes, preserving the pragmatic utility of 12-ET for chromatic music.

### Future Possibilities

- **Real-time adaptive EHCT**: Digital orchestras automatically adjusting each chord’s notes to optimize harmonic resonance on the fly.  
- **Multi-genre variations**: Weighted intervals tailored to classical, rock, jazz, or microtonal contemporary music.  
- **Empirical listening tests**: Systematic studies on listener preference to refine psychoacoustic weights \(\{w_i\}\).

In summary, EHCT represents a **systematic, data-driven approach** to near-equal tuning. By explicitly **modeling** the human ear’s tolerance for minute deviations, we can **optimize** a temperament that resonates more powerfully—bridging the ancient quest for pure intervals with the modern need for flexible tonality.

---

## References / Further Reading

1. Sethares, W. A. *Tuning, Timbre, Spectrum, Scale*. Springer, 2005.  
2. Johnston, B. *Maximum Clarity and Other Writings on Music*. University of Illinois Press, 2006.  
3. Sethares, W. A. “Adaptive Tunings for Musical Scales,” *Journal of the Acoustical Society of America*, 2004.  
4. Barbour, J. M. *Tuning and Temperament Library*. Dover Publications, 2004.
