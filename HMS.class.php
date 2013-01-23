<?php

class HMS {
    private $h;
    private $m;
    private $s;
    
    public function __construct($h, $m, $s) {
        $this->h = $h;
        $this->m = $m;
        $this->s = $s;
        return $this->getDecimal();
    }
    
    public function getDecimal() {
        return $this->h + ($this->m / 60) + ($this->s / 3600);
    }
    
    public function getDecimalDeg() {
        return ($this->h * 15) + ($this->m / 60) + ($this->s / 3600);
    }
}