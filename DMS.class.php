<?php

class DMS {
    private $d;
    private $m;
    private $s;
    
    public function __construct($d, $m, $s) {
        $this->d = $d;
        $this->m = $m;
        $this->s = $s;
        return $this->getDecimal();
    }
    
    public function getDecimal() {
        return $this->d + ($this->m / 60) + ($this->s / 3600);
    }
}