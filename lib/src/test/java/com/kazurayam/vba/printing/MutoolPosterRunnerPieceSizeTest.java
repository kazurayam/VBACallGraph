package com.kazurayam.vba.printing;

import org.testng.annotations.Test;

import static org.assertj.core.api.Assertions.assertThat;

public class MutoolPosterRunnerPieceSizeTest {

    @Test
    public void test_A4() {
        assertThat(MutoolPosterRunner.PieceSize.A4.getWidthMM()).isEqualTo(210);
        assertThat(MutoolPosterRunner.PieceSize.A4.getHeightMM()).isEqualTo(297);
    }

    @Test
    public void test_findByName() {
        assertThat(MutoolPosterRunner.PieceSize.findByName("A4"))
                .isEqualTo(MutoolPosterRunner.PieceSize.A4);
    }
}
