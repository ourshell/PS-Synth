# Near-Equal Temperament: A Complete Analysis

## Introduction

In acoustics, the accuracy of Equal Temperament (ET) has always been compromised, affecting the quality of sound. To address this, a new theory called **Near-Equal Temperament (N-ET)** has been developed. The core of N-ET is a stack of just intervals - 7 perfect fifths ($$\tfrac{3}{2}$$) and 1 major third ($$\tfrac{5}{4}$$). While the process of manual tuning is time-consuming, the mathematical precision of true Equal Temperament can be achieved with unconditional accuracy, a feat never accomplished before. This article explores the theory of N-ET through music theory, mathematics, and acoustics.

---

## 1. Sound Perception (Pitch)

The ability to distinguish between two frequencies is not linear. For example, increasing the frequency of $$A_4 = 440 \text{Hz}$$ by 50% results in $$660 \text{Hz}$$, but the perceived pitch increase is **not** proportional. This is due to the exponential relationship between pitch and frequency.

To increase the **perceived** pitch by 50% (a factor of $$\sqrt{2}$$), the frequency must be calculated as:

```math
\text{F} = 440 \times 2^{\frac{1}{2}} = 440 \times \sqrt{2} \approx 622 \text{Hz}.
```

In ET, frequencies are calculated to divide the octave into 12 equally spaced tones (as perceived by the human ear). The ratio for each tone is:

```math
\text{Ratio}_n = 2^{\tfrac{n}{12}}\, \quad n = 0, 1, \dots 11.
```

---

## 2. Linear Representation

Intervals in a musical scale are measured in **cents**, a logarithmic unit that linearizes pitch changes. The formula for pitch difference in cents between two frequencies $$f_1$$ and $$f_2$$ is:

```math
\text{Pitch} 
= 1200 \times \frac{\ln\Bigl(\tfrac{f_2}{f_1}\Bigr)}{\ln(2)} 
= 1200 \times \frac{\ln\Bigl(\tfrac{\text{Ratio}_2}{\text{Ratio}_1}\Bigr)}{\ln(2)} 
= 1200 \times \frac{\ln(\text{Ratio})}{\ln(2)}
```

For example, in ET, a semitone (100 cents) is calculated as:

```math
\text{Pitch}_{\text{semitone}} 
= 1200 \times \frac{\ln\Bigl(2^{\tfrac{1}{12}}\Bigr)}{\ln(2)} 
= 1200 \times \frac{0.05776226504}{0.69314718056} 
\approx 100 \text{cents}
```

Similarly, a whole tone (200 cents) between two tones in ET can be shown by:

```math
\text{Pitch}_{\text{whole-tone}} 
= 1200 \times \frac{\ln\Bigl(2^{\tfrac{7}{12}} \div 2^{\tfrac{5}{12}}\Bigr)}{\ln(2)} 
= 200 \text{cents}
```

---

## 3. Just Noticeable Difference (JND)

The **Just Noticeable Difference (JND)** defines the smallest pitch change perceivable by the human ear. In the range of 80 Hz to 500 Hz, JND is approximately **1 Hz** for complex sounds. For example:

```math
\text{JND}_1 = 1200 \times \frac{\ln\Bigl(\tfrac{500}{499}\Bigr)}{\ln(2)} \approx 3.465935 \text{cents}
```

For harmonic intervals, such as the [syntonic comma](https://en.wikipedia.org/wiki/Syntonic_comma) $$\bigl(\tfrac{81}{80}\bigr)$$, the difference is:

```math
\text{JND}_2 = 1200 \times \frac{\ln\Bigl(\tfrac{81}{80}\Bigr)}{\ln(2)} \approx 21.5063 \text{cents}
```

Pitch changes below **3.5 cents** are generally imperceptible, which becomes critical when comparing tuning systems.

---

## 4. Tuning Systems Comparison

### Perfect Fifths (P5)

In ET, the ratio of a perfect fifth is **approximated** as:

```math
\text{Ratio}_{\text{P5-ET}} = 2^{\tfrac{7}{12}} \approx 1.498307
```

In Just Intonation (JI), the perfect fifth is **exactly**:

```math
\text{Ratio}_{\text{P5-JI}} = \tfrac{3}{2} = 1.5
```

The difference in cents is:

```math
\text{Difference} = 1200 \times \frac{\ln\Bigl(\text{Ratio}_{\text{P5-ET}} \div \text{Ratio}_{\text{P5-JI}}\Bigr)}{\ln(2)}
= 1200 \times \frac{\ln\Bigl(2^{\tfrac{7}{12}} \div \tfrac{3}{2}\Bigr)}{\ln(2)} \approx -1.955 \text{cents}
```

Even a 2-cent discrepancy can contribute to subtle beating in chords.

---

## 5. Ratios Relationship

Within an octave, the most harmonic chord is the **major triad**, consisting of:

- Major third (M3)
- Perfect fifth (P5)

The deviation of M3 and P5 in ET compared to JI is calculated as follows:

For the major third:

```math
\text{Error}_{\text{M3}} = \Bigl[1200 \times \frac{\ln(5/4)}{\ln(2)}\Bigr] - 400 \approx -13.686 \text{cents}
```

For the perfect fifth:

```math
\text{Error}_{\text{P5}} = \Bigl[1200 \times \frac{\ln(3/2)}{\ln(2)}\Bigr] - 700 \approx +1.955 \text{cents}
```

These deviations are critical for harmonic purity.

### Stacking Intervals

By stacking **7 P5s** and **1 M3**, a near-perfect interval is achieved. The formula for the stack is:

```math
\text{Stack} = \Bigl(\tfrac{3}{2}\Bigr)^7 \times \Bigl(\tfrac{5}{4}\Bigr) \times \Bigl(\tfrac{1}{2^4}\Bigr).
```

Simplifying:

```math
\text{Stack} = \tfrac{10935}{8192} \approx 1.334839
```

Compared to ET:

```math
\text{Ratio}_{\text{P4, ET}} = 2^{\tfrac{5}{12}} \approx 1.334839
```

The error in cents is:

```math
\text{Error}_{\text{stack}} = 1200 \times \frac{\ln(1.334839)}{\ln(2)} - 500 \approx -0.0013 \text{cents}
```

---

## 6. Near-Equal Intervals

The stack can be **iteratively applied** to generate all 12 tones in the scale. The formula for each ratio is:

$$
\text{NER}_n 
= \Bigl(\tfrac{10935}{8192}\Bigr)^n, 
\quad n = 0, 1, \dots, 11.
$$

Ratios exceeding 2 are divided by 2 to stay within the octave. The worst cumulative error is:

$$
\text{Cumulative Error} 
= 11 \,\times\, (-0.0013) 
\;\approx\; -0.014 \,\text{cents}.
$$

---

## 7. Practical Applications

N-ET achieves **unprecedented precision**, with errors far below the JND threshold. This makes it suitable for:

1. **Acoustic instruments**, ensuring harmonic purity.  
2. **Digital synthesizers**, where computational precision can replicate N-ET ratios accurately.

---

## 8. Conclusion

Near-Equal Temperament offers a **revolutionary approach** to tuning, combining the harmonic purity of Just Intonation with the flexibility of Equal Temperament. By leveraging precise interval stacking, N-ET overcomes traditional tuning limitations and achieves near-perfect harmony. **Future research** could explore its application to microtonal systems and ensemble tuning, revolutionizing musical acoustics.

