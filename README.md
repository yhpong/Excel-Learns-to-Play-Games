# Excel-Learns-to-Play-Games
Reinforcement Learning with Proximal Policy Optimization

![MazeII](/Screenshots/MazeII_Solved.gif)

*This is more about me ranting than a technical discussion. If you are looking for resources to learning reinforcement learning, there are better resources out there, like [Lilian Weng's Blog](https://lilianweng.github.io/)*

I have long procrastinated from studying reinforcement learning. Partly because of busy work schedule, partly because the math looks intimidating and partly because there are already great minds out there who coded the libraries for us. But with all the recent advancements, I don't want to have too much to catch up in the next few years, so I decided to bite the bullet, and really learnt the nitty gritty, coded everything up from scratch, in VBA!! So no libraries, no vectorization, no auto-differentiation. Every matrix math is done in for-loops, all the gradients and back propagation are done by hand!

I decided to design some games and let Excel learn to play them. I started directly from policy gradient algorithm, since it sounded the most intuitive to me, at least on paper:
1. Start with a random policy, which is basically a neural network that takes environmental state variables as input, and outputs the probability of taking an action.
2. Use this random policy to let an agent intreact with the environment.
3. The environment would reward or penalize the agent's action depending on the task.
4. Repeat 2 & 3 until task is completed or time limit is reached.
5. Adjust the policy with gradient ascent in a way that would maximize rewards.
6. Repeat from step 2 with adjusted policy, until task can be completed with confidence.

While this plain vanilla policy gradient algorithm sounds good on paper, I quickly learnt from experience that it's almost always impractical. The problem is that every time we let the agent interacts with the environment and learns from that single experience, we don't know to what extend do we need to adjust its policy based on this particular experience. If we adjust it by too much, its next play may then give him a different experience that contradicts the preivous ones, and it ended up readjusting its policy everytime and things never converge. In a game with many possible paths to take, it may well be possible that he'd never even finish the task to gather the final reward, and thus failing to learn the task. So unless the task or game is defined in a very small finite set of states, in which we could play out all the possible paths, store them in memory, then the agent could directly learn from memory which of these paths are more rewarding. But we know that's impossible.

After much despair and thousands of epochs, I came across "Proximal Policy Optimization" (PPO) published by [Schulman et al in 2017](https://arxiv.org/abs/1707.06347v2). Great resources on how to actually implement them can be found in [ICRL Blog Track](https://iclr-blog-track.github.io/2022/03/25/ppo-implementation-details/). It basically tells you to what extend you can ajust the policy after gathering some experiences. With that implemented, my little games coded in Excel suddenly become achievable.

So let's get into action. I started with a toy problem which is adopted from https://www.samyzaf.com/ML/rl/qmaze.html. The game is basically a 7X7 grid maze which our hero (M) needs to traverse to reach the Princess (P), and our hero can only choose from 1 of 4 actions at any time step: move up, down, left or right.

![Maze1](/Screenshots/Maze_Features.png)

Let's start by giving our hero a brain, or a policy network, which is a simple feed forward neural network with 16 hidden units, connected to a softmax output layer of 4 dimensions, corresponding to the 4 selectable actions.

At first I followed the blog-post and trained our hero with a reward scheme like this:
1. Every time step carries a penalty.
2. Moving into an obstacle is penalized. Step is still incremented but player's position isn't changed.
3. Re-visiting an already visited position is penalized.
4. Successfully reaching the Princess is awarded.

This reward scheme makes a lot of sense, since it encourages the player to finish the game faster, and discourages it from meaningless actions like running into a wall and going back and forth. And obviously reaching its goal carries a reward. The environmetal states that were passed to the player were:
1. A 7X7 matrix with 1 indicating the player position, and 0 otherwise, flattened to a length 49 vector.

You can easily see why this is only a "toy" problem. Our hero knows in advance that the whole input space cannot be larger than 7X7. And he is basically blind to what's around him. Only by running into an obstacle and being punished for the action, would he learn to "not go right when my position is (1,7)". All he knows is that performing certain action at a certain location carries a penalty, and he learns to do less of it in his next play.

However, when I trained this blindfolded hero, it took a long time to make any meaningful progress (mind you this is VBA). So to speed up his training program, I added an incentive for him to explore the maze faster: 

5. Visiting a previously un-visited position is rewarded.

This is a complemetary reward to penalty #3, to encourage our hero to move forward instead of backward. The benefit of adding this reward is apparent:

![Maze_Solved](/Screenshots/Maze_Solved.gif)
![Maze_TrainProgress](/Screenshots/Maze_TrainingProgress.jpg)

With reward #5 added, he quickly managed to learn and reached 100% win rate pretty fast, in a total of 17 cycles, where each cycle he learnt from 50 episodes. One reason why he learnt to play this maze so quickly is that it's not really a maze. In this small maze, the obstacles often constrain our hero to move in specific directions. Since moving into obstacles and backtracking are both penalized, he quickly learns that going foward is more rewarding. So possibilites of going off target is limited.

Things however become much more challenging when I experimented with a bigger maze with a 7X13 grid as below:
![Larger Maze](/Screenshots/Maze_Compared.jpg)

The size of this maze is almost doubled, but the learning curve quickly becomes insurmountable. Difficulties soon arise when he enters the lower middle chamber. This is a huge void with no distinct feature to guide him. If he choose to enter the lower slot at the bottom right,  penalty #3 would encourage him to move deeper into the end, while at the same time he's penalized for moving back out, so that would be a local optimum for him to stay in. Even if he enters the middle area which is closer to the goal, it's a huge space with only a small window into where the Pricess is held captive. Using Monte Carlo to do this means only once in many episodes would he stand in front of that entry point, and many more episodes until he would actaully step through that window, and repeat that many times until he realizes the Princess is here. The animated gif below shows how our hero got stucked when he was trained to play this maze with the previous scheme:

![MazeII_stuck](/Screenshots/MazeII_stuck.gif)

If this is really a game and I am a gamer, this would be a pretty bad 1-star game design since the game makes you wander around without purpose, do a lot of back tracking, and the only way to win the game is by luck. However, this could well be a real problem that our hero really needs to solve. So to encourage it to learn faster, I adjusted both the reward scheme and the state variables that he was allowed to know. In the new reward scheme, the following change was made:

6. End-reward of reaching the Princess is increased.

By increasing the end reward was greatly increased so he didn't prefer exploration over the Princess (finishing all the side missions before the main plot). And if he was lucky enough to reach the Princess once, he would be very tempted to reach that goal again.

For the state variables, I also made him work more like a real person/robot, where he could see:
1. x & y coordinates of current position (rescaled to 1)
2. x & y coordinates of last position (rescaled to 1)
3. whether the 8 immediate cells surrouding him are blocked.
4. whether the 8 immediate cells surrouding him were previously visited.

So unlike last time where the whole 7X13 grid was encoded as states, now he only knows the immediate information surrounding him, just like how we would actually play a real maze. However the real game changer is state #4: he now has memory. He knows whether he has visited certain area already. By remembering whether there are un-explored spaces around him, combining this with reward #5, he would quikcly learn that exploration can pay off, which encourages him to step through that little window that leads to the Princess. And by increasing the reward of reaching the princess, it then quick learns to head straight to the princess without further ado.

![MazeII_progress](/Screenshots/MazeII_TrainingProgress_withvswithoutMemory.jpg)

In the figure above we can see that when our hero has memory, he easily solves the maze in about 15 cycles (750 games). Compared that to using the original training scheme, where he got stuck in a learning plateau and never even finished the game. So it's evident that the optimization algorithm is only one component here. How we designed the environment is just as, or even more important. In the real world, that would mean better sensors, better processing power, cleaner data and more common sense.

The takeaway from these experience is that reinforcement learning, or aritificial intelligence in general, is often exaggerated and misrepresented in general media. It is a magical thing to see some randomly intialized matricies to learn and adapt to tasks in the physical dimension, but at the same time they are not very different from linear regression: you fit a model of responses versus inputs. As such they fall into the same caveats, the training is as good as the training data you fed to it, and also the reward scheme that was set in the training program. On one hand, that means for any physicla tasks, we would need sensors that are capable of feeding the robot with good environmental variables. On the other hand, while some people would say AI models do not have human biases and are thus better suited for certain tasks, but are they? In the example above, what the model learnt was clearly affected by how we rewarded its actions and also what we allowed it to see. So it would certainly still carries the bias that the deisnger has when he/she designed the system. 

So would you trust a Robocop to go around handing out tickets and handcuffing criminals? For me the answer is no. Because at the end of the day these systems are very likely controlled by whatever authority who's already in power, and that Robocop would certainly be tuned to do the job "perfectly" for the benefit of its designers. And history tells us that this never ends well. In a way, one could argue that this is not very different from the system we have today. Our legal and justice system is designed by those in power, and those who enforce it are encourged to do law's bidding and crush those who oppose it. But if one believes in free will (let's not enter the free will debate for now) and the capacity of man to act out of compassion and justice, then there are chances that our system will at some point head to a brighter state.
