#       Base  Plus  Minus
# Category 1  0 4.3 0 0
# Category 2  4.3 0 4.4 0
# Category 3  6.9 0 0 1.8
# Category 4  4.1 0 0 2.8
# Category 5  4.1 0 3.2 0
# Category 6  0 7.3 0 0

# bar = {
#   'Elephant' => [rand(1..10), rand(1..10), rand(1..10)],
#   'Fudge' => [rand(1..10), rand(1..10), rand(1..10)],
#   'Google' => [rand(1..10), rand(1..10), rand(1..10)],
#   'Hippo' => [rand(1..10), rand(1..10), rand(1..10)]
# }

input = [4.3, 4.4, -1.8, -2.8, 3.2]
output = {
  :fill => [],
  :base => [],
  :plus => [],
  :minu => []
}

input << 0
input.each_with_index do |num, index|
  if index == 0
    output[:fill] << 0
    output[:base] << num
    output[:plus] << 0
    output[:minu] << 0
  elsif index == input.length - 1
    output[:base] << output[:fill].last + output[:plus].last if output[:plus].last > 0
    output[:base] << output[:fill].last if output[:minu].last < 0
    output[:fill] << 0
    output[:plus] << 0
    output[:minu] << 0
  elsif num > 0
    output[:fill] << (output[:fill].last + output[:base].last + output[:plus].last)
    output[:base] << 0
    output[:plus] << num
    output[:minu] << 0
  elsif num < 0
    output[:fill] << (output[:fill].last + output[:base].last + output[:plus].last + num)
    output[:base] << 0
    output[:plus] << 0
    output[:minu] << num
  end
end



puts output
