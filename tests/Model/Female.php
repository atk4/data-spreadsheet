<?php
declare(strict_types=1);

namespace atk4\data\tests\Model;

class Female extends Person
{
    public function init()
    {
        parent::init();
        $this->addCondition('gender', 'F');
    }
}
