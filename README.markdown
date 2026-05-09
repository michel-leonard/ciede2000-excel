Browse : [Java](https://github.com/michel-leonard/ciede2000-java) · [Julia](https://github.com/michel-leonard/ciede2000-julia) · [Kotlin](https://github.com/michel-leonard/ciede2000-kotlin) · [Lua](https://github.com/michel-leonard/ciede2000-lua) · [MATLAB](https://github.com/michel-leonard/ciede2000-matlab) · **Microsoft Excel** · [PHP](https://github.com/michel-leonard/ciede2000-php) · [Perl](https://github.com/michel-leonard/ciede2000-perl) · [Python](https://github.com/michel-leonard/ciede2000-python) · [R](https://github.com/michel-leonard/ciede2000-r) · [Ruby](https://github.com/michel-leonard/ciede2000-ruby)

# CIEDE2000 color difference formula in Excel

This page presents the CIEDE2000 color difference, implemented in Microsoft Excel.

![Logo](https://raw.githubusercontent.com/michel-leonard/ciede2000-color-matching/refs/heads/main/docs/assets/images/logo.jpg)

## Our CIEDE2000 offer

This production-ready file, released in 2026, contain the CIEDE2000 algorithm.

Source File|Type|Bits|Purpose|Advantage|
|:--:|:--:|:--:|:--:|:--:|
[ciede2000.xls](./ciede2000.xls)|`Number`|64|General|Pure native formula, with no macros or [VBA](https://github.com/michel-leonard/ciede2000-vba)|

Tested by Michel LEONARD on 50,000 carefully chosen pairs of colors, with tests that also take into account the parametric factors `kL`, `kC` and `kH`, this Excel formula provides at least 11 accurate decimal places.

<details>
<summary>Do you have a LAMBDA CIEDE2000 feature for Microsoft Excel Online users?</summary>

### LAMBDA

We’ve created a LAMBDA function for Excel 365 that will let you easily compare all your color pairs.

```excel
=LAMBDA(l_1, a_1, b_1, l_2, a_2, b_2, k_l, k_c, k_h, canonical,
    LET(
        comment, "CIEDE2000 function for Excel 365 users",
        pi, PI(),
        b_1_sq, b_1 * b_1,
        b_2_sq, b_2 * b_2,
        g_factor, LET(
            comment, "Compute G compensation factor",
            c_1_orig, SQRT(a_1 * a_1 + b_1_sq),
            c_2_orig, SQRT(a_2 * a_2 + b_2_sq),
            c_avg, 0.5 * (c_1_orig + c_2_orig),
            c_avg_7, POWER(c_avg, 7),
            1.5 - 0.5 * SQRT(c_avg_7 / (c_avg_7 + 6103515625))
        ),
        a_1_prime, a_1 * g_factor,
        a_2_prime, a_2 * g_factor,
        c_1, SQRT(a_1_prime * a_1_prime + b_1_sq),
        c_2, SQRT(a_2_prime * a_2_prime + b_2_sq),
        h_1, LET(
            comment, "Compute hue angles in radians, and adjust for negative",
            h_1_raw, IF((a_1 = 0) * (b_1 = 0), 0, ATAN2(a_1_prime, b_1)),
            IF(h_1_raw < 0, h_1_raw + 2 * pi, h_1_raw)
        ),
        h_2, LET(
            h_2_raw, IF((a_2 = 0) * (b_2 = 0), 0, ATAN2(a_2_prime, b_2)),
            IF(h_2_raw < 0, h_2_raw + 2 * pi, h_2_raw)
        ),
        pi_interoperability, pi + 0.00000000000001,
        note_1, "If you replace the constant 0.00000000000001 by zero, the code",
        note_2, "will continue to work, but CIEDE2000 interoperability between",
        note_3, "all programming languages will no longer be guaranteed",
        cond, pi_interoperability < ABS(h_2 - h_1),
        h_mean, LET(
            hm, 0.5 * (h_1 + h_2),
            comment, "Hue mean wraps around pi (180 deg)",
            IF(cond,
                IF(canonical * (pi_interoperability < hm),
                    hm - pi,
                    hm + pi
                ), hm
            )
        ),
        h_delta, LET(
            hd, 0.5 * (h_2 - h_1),
            IF(cond, hd + pi, hd)
        ),
        c_mean, 0.5 * (c_1 + c_2),
        r_t, LET(
            comment, "Compute hue rotation correction factor R_T",
            c_mean_7, POWER(c_mean, 7),
            r_c_sq, c_mean_7 / (c_mean_7 + 6103515625),
            r_c, SQRT(r_c_sq),
            theta_num, 36 * h_mean - 55 * pi,
            theta, theta_num / (5 * pi),
            theta_sq, POWER(theta, 2),
            -2 * r_c * SIN(pi / 3 * EXP(-theta_sq))
        ),
        l_term, LET(
            comment, "Compute lightness term",
            l_mean, 0.5 * (l_1 + l_2),
            l_mean_sq, POWER(l_mean - 50, 2),
            s_l_num, 0.015 * l_mean_sq,
            s_l_denom, SQRT(20 + l_mean_sq),
            s_l, 1 + s_l_num / s_l_denom,
            l_num, l_2 - l_1,
            l_num / (k_l * s_l)
        ),
        c_term, LET(
            comment, "Compute chroma term",
            s_c, 1 + 0.045 * c_mean,
            c_num, c_2 - c_1,
            c_num / (k_c * s_c)
        ),
        h_term, LET(
            comment, "Compute hue term",
            trig_1, 0.17 * SIN(h_mean + pi / 3),
            trig_2, 0.24 * SIN(2 * h_mean + 0.5 * pi),
            trig_3, 0.32 * SIN(3 * h_mean + 8 * pi / 15),
            trig_4, 0.2 * SIN(4 * h_mean + 3 * pi / 20),
            trig, 1 - trig_1 + trig_2 + trig_3 - trig_4,
            s_h, 1 + 0.015 * trig * c_mean,
            h_num, 2 * SQRT(c_1 * c_2) * SIN(h_delta),
            h_num / (k_h * s_h)
        ),
        delta_e, LET(
            comment, "Combine lightness, chroma, hue, and interaction terms",
            l_part, l_term * l_term,
            c_part, c_term * c_term,
            h_part, h_term * h_term,
            interaction, c_term * h_term * r_t,
            SQRT(l_part + c_part + h_part + interaction)
        ),
        delta_e
    )
)
```

This LAMBDA function is interoperable with our other CIEDE2000 functions, works with both scalars and vectors, and as such provides 11 correct decimal places. Lead developer LEONARD has successfully tested it on 1,000,000 well-chosen L\*a\*b\* color pairs, and with various parametric factors `kL`, `kC` and `kH`.

### Example Usage

Let’s assume that Excel columns contain `L1,a1,b1,L2,a2,b2,kL,kC,kH`, then you can replace `@@@` by your expression `=LAMBDA(...)(A1:A3,B1:B3,C1:C3,D1:D3,E1:E3,F1:F3,G1:G3,H1:H3,I1:I3,TRUE)`:

```text
78.7	65.2	-2.9	77.5	60.7	2.8	1.0	1.0	1.0	@@@
88.3	126.1	-1.7	89.3	109.1	4.6	1.0	1.0	0.9
38.3	112.8	-19.9	28.7	35.1	6.4	1.1	1.0	0.9
```

This will directly populate all the CIEDE2000 color differences, calculated vectorized by the LAMBDA:

```text
78.7	65.2	-2.9	77.5	60.7	2.8	1.0	1.0	1.0	2.9198852526166
88.3	126.1	-1.7	89.3	109.1	4.6	1.0	1.0	0.9	3.5257373508076
38.3	112.8	-19.9	28.7	35.1	6.4	1.1	1.0	0.9	21.8114219143342
```

The last parameter `TRUE` ensures that your results are consistent with those calculated by [Gaurav Sharma](https://hajim.rochester.edu/ece/sites/gsharma/ciede2000/) in MATLAB.

<details>
<summary>How can I simply leave the last 4 parameters at their default values?</summary>

```excel
=LAMBDA(l_1, a_1, b_1, l_2, a_2, b_2,
    LET(
        ciede2000_with_parameters, LAMBDA(...),
        ciede2000_with_parameters(l_1, a_1, b_1, l_2, a_2, b_2, 1, 1, 1, FALSE)
    )
)
```

This will allow you to call the LAMBDA function with just 6 parameters, representing your two L\*a\*b\* colors.

</details>

</details>

### Software Versions

- Microsoft Excel 97-2003
- Microsoft 365 | Excel
- Google Sheets 2026

## Public Domain Licence

You are free to use these files, even for commercial purposes.
