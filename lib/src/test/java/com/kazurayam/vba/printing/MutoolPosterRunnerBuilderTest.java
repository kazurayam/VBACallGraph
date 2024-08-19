package com.kazurayam.vba.printing;

import com.kazurayam.vba.printing.MutoolPosterRunner.Builder;
import org.testng.annotations.Test;

import static org.assertj.core.api.Assertions.assertThat;

public class MutoolPosterRunnerBuilderTest {

    @Test
    public void test_resolveDecimationFactor() {
        Builder builder = new MutoolPosterRunner.Builder();
        // 210 millimeter is the width of A4 pater
        assertThat(builder.deriveDecimationFactor(200, 210)).isEqualTo(1);
        assertThat(builder.deriveDecimationFactor(210, 210)).isEqualTo(1);
        assertThat(builder.deriveDecimationFactor(211, 210)).isEqualTo(2);
        assertThat(builder.deriveDecimationFactor(420, 210)).isEqualTo(2);
        assertThat(builder.deriveDecimationFactor(421, 210)).isEqualTo(3);
        assertThat(builder.deriveDecimationFactor(630, 210)).isEqualTo(3);
        assertThat(builder.deriveDecimationFactor(631, 210)).isEqualTo(4);
    }
}
