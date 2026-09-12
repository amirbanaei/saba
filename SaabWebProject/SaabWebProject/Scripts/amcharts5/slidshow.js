<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script>
        let timer;
        let slides = ["#slide1", "#slide2", "#slide3"];
        let currentSlide = 0;

        function showSlides() {
            $(slides[currentSlide]).addClass("show-slide");
            setTimeout(function() {
            $(slides[currentSlide]).removeClass("show-slide");
        currentSlide++;
                if (currentSlide >= slides.length) {
            currentSlide = 0;
    }
    showSlides();
}, 5000); // Change this value to adjust slide duration
}

        function hideSlides() {
            for (let i = 0; i < slides.length; i++) {
            $(slides[i]).removeClass("show-slide");
    }
}

// Hide the slides when the page first loads
hideSlides();

// Set up the mousemove event listener
        $(document).mousemove(function() {
            clearTimeout(timer);
        hideSlides();
        timer = setTimeout(showSlides, 10 * 1000); // Set timer for 10 seconds
    });

    // Set timer to show slides after 10 seconds
    timer = setTimeout(showSlides, 10 * 1000);
    </script>